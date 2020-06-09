<%
  Dim vJobs_AcctId, vJobs_Id, vJobs_Level, vJobs_Type, vJobs_Title, vJobs_Mods, vJobs_Ratings, vJobs_SkillMods, vJobs_Active
  Dim vJobs_Eof, vJobsListCnt, vJobsListMax

  '...Get all Jobs
  Sub sGetJobs_Rs
    vSql = "SELECT * FROM Jobs WHERE Jobs_AcctId = '" & svCustAcctId & "' ORDER BY Jobs_Id"
'   sDebug
    sOpenDb3    
    Set oRs3 = oDb3.Execute(vSql)
  End Sub


  '...same as above except for order
  Sub sGetJobs_Rs_ByTitle 
    vSql = "SELECT * FROM Jobs WHERE Jobs_AcctId = '" & svCustAcctId & "' ORDER BY Jobs_Title"
'   sDebug
    sOpenDb3    
    Set oRs3 = oDb3.Execute(vSql)
  End Sub


  '...Get Jobs Record
  Sub sGetJobs  (vJobsId)
    vJobs_Eof = False
    vSql = "SELECT * FROM Jobs WHERE Jobs_Id = '" & vJobsId & "' AND Jobs_AcctId = '" & svCustAcctId & "'"
'   sDebug
    sOpenDb3    
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then 
      sReadJobs
      vJobs_Eof = True
    End If
    Set oRs3 = Nothing
    sCloseDb3
  End Sub

  '...Get Jobs Type (Mandatory/Optional)
  Function fJobsType (vJobsId)
    fJobsType = ""
    If Len(vJobs_Id) > 0 Then 
      vSql = "SELECT Jobs_Type FROM Jobs WHERE Jobs_Id = '" & vJobsId & "'"
      sOpenDb3    
      Set oRs3 = oDb3.Execute(vSql)
      If Not oRs3.Eof Then 
        fJobsType = oRs3("Jobs_Type")
      End If
      Set oRs3 = Nothing
      sCloseDb3    
    End If
  End Function


  '...Get Jobs Title
  Function fJobsTitle (vJobsId)
    fJobsTitle = ""
    If Len(vJobsId) > 0 Then 
      vSql = "SELECT Jobs_Title FROM Jobs WHERE Jobs_Id = '" & vJobsId & "'"
      sOpenDb3    
      Set oRs3 = oDb3.Execute(vSql)
      If Not oRs3.Eof Then 
        fJobsTitle = oRs3("Jobs_Title")
      End If
      Set oRs3 = Nothing
      sCloseDb3    
    End If
  End Function
  
  
  '...Is Job Active>
  Function fJobsActive (vJobsId)
    fJobsActive = False
    If Len(vJobsId) > 0 Then 
      vSql = "SELECT Jobs_Active FROM Jobs WHERE Jobs_Id = '" & vJobsId & "'"
      sOpenDb3    
      Set oRs3 = oDb3.Execute(vSql)
      If Not oRs3.Eof Then 
        fJobsActive = oRs3("Jobs_Active")
      End If
      Set oRs3 = Nothing
      sCloseDb3    
    End If
  End Function
  

  
  Sub sReadJobs 
  '...same as below but it does NOT convert the XX in the mods so 
  '   you can see what's in the table
    vJobs_AcctId       = oRs3("Jobs_AcctId")
    vJobs_Id           = oRs3("Jobs_Id")
    vJobs_Level        = oRs3("Jobs_Level")
    vJobs_Type         = oRs3("Jobs_Type")
    vJobs_Title        = oRs3("Jobs_Title")
    vJobs_Mods         = oRs3("Jobs_Mods")
    vJobs_Ratings      = oRs3("Jobs_Ratings")
    vJobs_SkillMods    = oRs3("Jobs_SkillMods")
    vJobs_Active       = oRs3("Jobs_Active")
  End Sub


  Sub sReadJobsXX
  '...same as above convert the XX in the mods to 
  '   session language
    vJobs_AcctId       = oRs3("Jobs_AcctId")
    vJobs_Id           = oRs3("Jobs_Id")
    vJobs_Level        = oRs3("Jobs_Level")
    vJobs_Type         = oRs3("Jobs_Type")
    vJobs_Title        = oRs3("Jobs_Title")
    vJobs_Mods         = oRs3("Jobs_Mods") '...this was added in Jan 2008 to allow for multi-lingual jobs/progs
    vJobs_Mods         = Replace(Ucase(vJobs_Mods), "XX", svLang)
    vJobs_Ratings      = oRs3("Jobs_Ratings")
    vJobs_SkillMods    = oRs3("Jobs_SkillMods")
    vJobs_Active       = oRs3("Jobs_Active")
  End Sub


  Sub sExtractJobs
    vJobs_Id           = fNoQuote(Request.Form("vJobs_Id"))
    vJobs_Level        = Request.Form("vJobs_Level")
    vJobs_Type         = Request.Form("vJobs_Type")
    vJobs_Title        = fUnQuote(Request.Form("vJobs_Title"))
    vSkillNo           = Request.Form("vSkillNo")
    vJobs_Mods         = fUnQuote(Request.Form("vJobs_Mods"))
    vJobs_Active       = Request.Form("vJobs_Active")

    '...if using skills, get ratings and mods 
    If vSkillNo >= 1 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_1") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_1") & "~" 
    If vSkillNo >= 2 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_2") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_2") & "~" 
    If vSkillNo >= 3 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_3") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_3") & "~" 
    If vSkillNo >= 4 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_4") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_4") & "~" 
    If vSkillNo >= 5 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_5") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_5") & "~" 
    If vSkillNo >= 6 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_6") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_6") & "~" 
    If vSkillNo >= 7 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_7") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_7") & "~" 
    If vSkillNo >= 8 Then vJobs_Ratings = vJobs_Ratings & Request.Form("vRate_8") & "~" : vJobs_SkillMods = vJobs_SkillMods & Request.Form("vMods_8") & "~" 
    
    '...strip off trialing "~"     
    If vSkillNo > 0 Then
      vJobs_Ratings   = Left(vJobs_Ratings, Len(vJobs_Ratings)-1)
      vJobs_SkillMods = Trim(fNoQuote(Left(vJobs_SkillMods, Len(vJobs_SkillMods)-1)))
    End If

  End Sub
  
  Sub sInsertJobs
    vSql = "INSERT INTO Jobs "
    vSql = vSql & "(Jobs_AcctId, Jobs_Id, Jobs_Level, Jobs_Type, Jobs_Title, Jobs_Mods, Jobs_Ratings, Jobs_SkillMods, Jobs_Active)"
    vSql = vSql & " VALUES ('" & svCustAcctId & "', '" & vJobs_Id & "', '" & vJobs_Level & "', '" & vJobs_Type & "', '" & vJobs_Title & "', '" & vJobs_Mods & "', '" & vJobs_Ratings & "', '" & vJobs_SkillMods & "', " &  fSqlBoolean (vJobs_Active) & ")"
'   sDebug
    On Error Resume Next
    vFileOK = False   
    sOpenDb3
    oDb3.Execute(vSql)
    If Err.Number = 0 or Err.Number = "" Then 
      vFileOk = True
    Else
      vFileDesc = Err.Description
    End If
    On Error GoTo 0
    sCloseDb3
  End Sub

  Sub sUpdateJobs
    vSql = "UPDATE Jobs SET"
    vSql = vSql & " Jobs_Type      = '" & vJobs_Type      & "', " 
    vSql = vSql & " Jobs_Level     = '" & vJobs_Level     & "', " 
    vSql = vSql & " Jobs_Title     = '" & vJobs_Title     & "', " 
    vSql = vSql & " Jobs_Mods      = '" & vJobs_Mods      & "', " 
    vSql = vSql & " Jobs_Ratings   = '" & vJobs_Ratings   & "', " 
    vSql = vSql & " Jobs_SkillMods = '" & vJobs_SkillMods & "', " 
    vSql = vSql & " Jobs_Active    =  " & fSqlBoolean (vJobs_Active)
    vSql = vSql & " WHERE Jobs_Id  = '" & vJobs_Id & "' AND Jobs_AcctId = '" & svCustAcctId & "'"
    sOpenDb3
'   sDebug
    oDb3.Execute(vSql)
    sCloseDb3
  End Sub
  
  Sub sDeleteJobs
    vSql = "DELETE FROM Jobs WHERE Jobs_Id = '" & vJobs_Id & "' And Jobs_AcctId = '" & svCustAcctId & "'"
    sOpenDb3
    oDb3.Execute(vSql)
    sCloseDb3
  End Sub

  '...get all Jobs (Dropdown)
  Function fJobsOptions (vId)
    Dim vSelected
    fJobsOptions = ""
    vJobsListCnt = 0
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      vSelected = "" 
      If Instr(vId, oRs3("Jobs_Id")) > 0 Then 
        vSelected = " selected" 
      End If
      fJobsOptions = fJobsOptions & "<option value=" & Chr(34) & oRs3("Jobs_Id") & Chr(34) & vSelected & ">" & oRs3("Jobs_Id") & " - " & oRs3("Jobs_Title") & "</option>" & vbCrLf
      If vJobsListCnt < 8 Then
        vJobsListCnt = vJobsListCnt + 1
      End If      
      oRs3.MoveNext
    Loop      
    sCloseDb3           
    Set oRs3 = Nothing
  End Function


  '...get all Jobs|Type|Progs (Dropdown)
  Function fJobsProgOptions (vId)
    Dim j, vSelected, aMods
    fJobsProgOptions = ""
    vJobsListCnt = 1
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      If vJobs_Active Then
        aMods = Split(Trim(vJobs_Mods), " ")
        For j = 0 to Ubound(aMods)               
          vSelected = "" 
          If Instr(vId, vJobs_Id & "|" & vJobs_Type & "|" & aMods(j)) > 0  Then 
            vSelected = " selected" 
          End If
          If vJobsListCnt < 9 Then vJobsListCnt = vJobsListCnt + 1
          fJobsProgOptions = fJobsProgOptions & "<option value=" & Chr(34) & vJobs_Id & "|" & vJobs_Type & "|" & aMods(j) & Chr(34) & vSelected & ">" & vJobs_Id & " | " & aMods(j) & "</option>" & vbCrLf
        Next
      End If
      oRs3.MoveNext
    Loop      
    sCloseDb3           
    Set oRs3 = Nothing
  End Function
  
  
  '...get all Jobs|Type|Prog|Mods (Dropdown)   - Note: Requires Db_Prog.asp include
  Function fJobsModsOptions (vId)
    Dim i, j, vSelected, aProgs, aMods
    fJobsModsOptions = ""
    vJobsListCnt = 1
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      If vJobs_Active Then
        aProgs = Split(Trim(vJobs_Mods), " ")
        For i = 0 to Ubound(aProgs)               
          aMods = Split(fProgMods(aProgs(i)), " ")
          For j = 0 To Ubound(aMods)
            vSelected = "" 
            If Instr(vId, vJobs_Id & "|" & vJobs_Type & "|" & aProgs(i) & "|" & aMods(j)) > 0  Then 
              vSelected = " selected" 
            End If  
            If vJobsListCnt < 12 Then vJobsListCnt = vJobsListCnt + 1 '...this determines the depth of the dropdown menu (max 8 lines)
            fJobsModsOptions = fJobsModsOptions & "<option value=" & Chr(34) & vJobs_Id & "|" & vJobs_Type & "|" & aProgs(i) & "|" & aMods(j) & Chr(34) & vSelected & ">" & vJobs_Id & " | " & aProgs(i) & " | " & aMods(j) & "&nbsp;&nbsp;</option>" & vbCrLf            
          Next            
        Next
      End If
      oRs3.MoveNext
    Loop      
    sCloseDb3           
    Set oRs3 = Nothing
  End Function  
  

  '...display jobs in user.asp if they are not on the criteria talbe and if some or all jobs are not marketd "M" mandatory (meaning no selection)
  Function fDisplayJobs ()
    Dim vCountAll, vCountMandatory
    vCountAll = 0
    vCountMandatory = 0    
    sOpenDb3
    vSql = "SELECT COUNT(Jobs.Jobs_Type) AS [Count_All] " _
         & "FROM "_
         & "  Crit INNER JOIN Jobs ON CHARINDEX(Jobs.Jobs_Id, Crit.Crit_JobsId) > 0 " _
         & "WHERE "_ 
         & "  (Crit.Crit_AcctId = '" & svCustAcctId & "') AND "_
         & "  (Jobs.Jobs_AcctId = '" & svCustAcctId & "')"
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then vCountAll = oRs3("Count_All")
    vSql = "SELECT COUNT(Jobs.Jobs_Type) AS [Count_Mandatory] " _
         & "FROM Crit INNER JOIN Jobs ON Crit.Crit_JobsId = Jobs.Jobs_Id " _
         & "WHERE (Crit.Crit_AcctId = '" & svCustAcctId & "') AND (Jobs.Jobs_AcctId = '" & svCustAcctId & "')" _
         & "AND (Jobs.Jobs_Type = 'M')"
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then vCountMandatory = oRs3("Count_Mandatory")  
    fDisplayJobs = True
    If vCountAll > 0 Then 
      If vCountAll = vCountMandatory Then
        fDisplayJobs = False
      End If
    End If
  End Function



  '...get all Active Jobs|Progs - called by User.asp (and used in My Learning)
  Function fJobsProgsOptions (vId)
    Dim j, k, vSpaces, vSelected, aMods, vProgs
    fJobsProgsOptions = ""
    vJobsListCnt = 0
    vJobsListMax = 0

    sGetJobs_Rs
    '...get largest title size for optics on dropdown
    Do While Not oRs3.Eof 
      sReadJobs
      vJobsListMax = fMax(vJobsListMax, Len(vJobs_Title))
      oRs3.MoveNext
    Loop      

    sCloseDb3           
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      vJobsListMax = fMax(vJobsListMax, Len(vJobs_Title))
      If vJobs_Active Then
        vProgs = ""
        aMods = Split(Trim(vJobs_Mods), " ")
        For j = 0 to Ubound(aMods)               
          vSelected = "" : If Instr(vId, vJobs_Id) > 0  Then vSelected = " selected" 
          vProgs = vProgs & aMods(j) & " " 
        Next
        vProgs = Trim(vProgs)
        If vJobsListCnt < 8 Then vJobsListCnt = vJobsListCnt + 1
        vSpaces = ""
        For k = 1 to vJobsListMax - Len(Replace(vJobs_Title, "  ", " ")) + 3
          vSpaces = vSpaces & "&nbsp;"
        Next
        
        
        fJobsProgsOptions = fJobsProgsOptions  & "<option value=" & Chr(34) & vJobs_Id & "|" & vProgs & Chr(34) & vSelected & ">" & vJobs_Title & vSpaces & vJobs_Id & " (" & vProgs & ")</option>" & vbCrLf
      End If
      oRs3.MoveNext
    Loop      
    sCloseDb3           
    Set oRs3 = Nothing
  End Function



  '...get all Jobs|Progs (for Lao report) - must be active
  '   further reduce if vPrograms <> "" (J0001EN|M|P1234EN,J0002EN|M|P2134EN)
  Function fJobsProgAll (vPrograms)
    Dim j, aMods
    fJobsProgAll = ""
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      '...get job id and append programs for those required 
      If vJobs_Active Then
        aMods = Split(Trim(vJobs_Mods), " ")
        For j = 0 to Ubound(aMods)               
          '...before we add this to the string, check against vPrograms
          If vPrograms = "A" Or Instr(vPrograms, aMods(j)) > 0 Then
            fJobsProgAll = fJobsProgAll & vJobs_Id & "|" & vJobs_Type & "|" & aMods(j) & " "
          End If
        Next
      End If            
      oRs3.MoveNext
    Loop      
    sCloseDb3
    Set oRs3 = Nothing
    fJobsProgAll = Trim(fJobsProgAll) '...remove trailing space
  End Function


  '...get all active Jobs|Progs|Mods for Lao report
  '   reduce if vPrograms <> "" (J0001EN|M|P1234EN,J0002EN|M|P2134EN)
  Function fJobsProgModsAll (vPrograms)
    Dim i, j, aProgs, aMods
    fJobsProgModsAll = ""
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      '...get job id and append programs for those required 
      If vJobs_Active Then
        aProgs = Split(Trim(vJobs_Mods), " ")
        For i = 0 to Ubound(aProgs)
          aMods = Split(fProgMods(aProgs(i)), " ")
          For j = 0 to Ubound(aMods)
            '...before we add this to the string, check against vPrograms
            If vPrograms = "A" Or Instr(vPrograms, aProgs(i) & "|" & aMods(j)) > 0 Then
              fJobsProgModsAll = fJobsProgModsAll & vJobs_Id & "|" & vJobs_Type & "|" & aProgs(i) & "|" & aMods(j) & " "
            End If          
          Next
        Next
      End If            
      oRs3.MoveNext
    Loop      
    sCloseDb3
    Set oRs3 = Nothing
    fJobsProgModsAll = Trim(fJobsProgModsAll) '...remove trailing space
  End Function


  '...remove A Jobs|Progs (for Lao report) not on member record
  Function fJobsProg (vJobsAll, vJobsId)
    Dim i, aMods
    aMods = Split(vJobsAll)
    For i = 0 to Ubound(aMods)               
      If Mid(aMods(i), 9, 1) <> "A"  Or (Mid(aMods(i), 9, 1) = "A" And Instr(vJobsId, Left(aMods(i),7)) > 0) Then
        fJobsProg = fJobsProg & aMods(i) & " "
      End If
    Next
    fJobsProg = Trim(fJobsProg) '...remove trailing space
  End Function


  '...get all Jobs|Progs (for Lao report) - must be active
  '   extract subset if vJobsId   <> "" (J0001EN J0002EN)
  '   further reduce if vPrograms <> "" (J0001EN|M|P1234EN,J0002EN|M|P2134EN)
  Function fJobsProg_Original (vJobsId, vPrograms)
    Dim j, aMods
    fJobsProg = ""
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      '...get job id and append programs for those required 
      If vJobs_Active And (vJobs_Type = "M" or vJobs_Type = "O" Or vJobsId = "" Or Instr(vJobsId, vJobs_Id) > 0) Then
        aMods = Split(Trim(vJobs_Mods), " ")
        For j = 0 to Ubound(aMods)               
          '...before we add this to the string, check against vPrograms
          If vPrograms = "A" Or Instr(vPrograms, aMods(j)) > 0 Then
            fJobsProg = fJobsProg & vJobs_Id & "|" & vJobs_Type & "|" & aMods(j) & " "
          End If
        Next
      End If            
      oRs3.MoveNext
    Loop      
    sCloseDb3
    Set oRs3 = Nothing
    fJobsProg = Trim(fJobsProg) '...remove trailing space
  End Function
 
  
  '...Get Active Jobs record for Active Member (very sexy sql here)
  '   note, if user has multi criteria, there could be more than one job record
  Sub sGetJobsByMemb
    Dim i, aMemb_Crit, vProgs
    vJobs_Mods = ""
    vJobs_Eof = True
    '...no jobs programs can be assigned unless there's a non zero criteria
    If svMembCriteria = "0" Then Exit Sub
    aMemb_Crit = Split(svMembCriteria)
    '...get any job records assigned to criteria in this account
    vSql = "SELECT Crit.Crit_No, Jobs.Jobs_Mods, Jobs.Jobs_Level  " _
         & "FROM Crit INNER JOIN " _
         & "Jobs ON CHARINDEX(Jobs.Jobs_Id, Crit.Crit_JobsId) > 0 "_
         & "WHERE (Crit.Crit_AcctId = '" & svCustAcctId & "') " _
         & "AND (Jobs.Jobs_AcctId = '" & svCustAcctId & "') " _
         & "AND (Jobs.Jobs_Active = 1) "
'   sDebug
    sOpenDb3
    Set oRs3 = oDb3.Execute(vSql)   
    Do While Not oRs3.Eof 
      '...get jobs if ok for this learner level
      If oRs3("Jobs_Level") = "A" Or (oRs3("Jobs_Level") = "L" and svMembLevel = 2) Or (oRs3("Jobs_Level") = "F" and svMembLevel > 2) Then
        For i = 0 To Ubound(aMemb_Crit)
          If Cstr(oRs3("Crit_No")) = aMemb_Crit(i) Then
            '...this was added in Jan 2008 to allow for multi-lingual jobs/progs
            vProgs = Replace(Ucase(oRs3("Jobs_Mods")), "XX", svLang)
            vJobs_Mods = vJobs_Mods + " " & vProgs    '...add a space between strings
          End If
        Next    
      End If
      oRs3.MoveNext
    Loop    
    vJobs_Mods = Trim(vJobs_Mods) '...remove extra spaces      
    sCloseDb3
    Set oRs3 = Nothing
  End Sub
  
  '...get the next available Jobs ID for this account (go down or up)
  Function fNextJobsId ()
    vSql = "SELECT "_
         & " MIN(CAST(SUBSTRING(Jobs_Id, 2, 4) AS int)) AS JobsMin, "_
         & " MAX(CAST(SUBSTRING(Jobs_Id, 2, 4) AS int)) AS JobsMax "_
         & "FROM [V5_Vubz].[dbo].[Jobs] WITH (nolock) "_
         & "WHERE Jobs_AcctId = '" & svCustAcctId & "'"
         
    sOpenDb3    
    Set oRs3 = oDb3.Execute(vSql)
    '...start with smallest value and subtract 1 (unless already at 1)
    If oRs3("JobsMin") > 1 Then
      fNextJobsId = "J" & Right("0000" & oRs3("JobsMin") - 1, 4) & svLang
    ElseIf oRs3("JobsMax") < 9999 Then
      fNextJobsId = "J" & Right("0000" & oRs3("JobsMax") + 1, 4) & svLang
    Else
      fNextJobsId = ""
    End If
    Set oRs3 = Nothing
    sCloseDb3
  End Function  
   

  
%>