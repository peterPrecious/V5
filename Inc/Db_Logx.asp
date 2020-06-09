<%
  Dim vLogx_No, vLogx_MembNo, vLogx_ProgNo, vLogx_ModsNo, vLogx_Expires, vLogx_BestDate, vLogx_BestScore, vLogx_NoAttempts, vLogx_TimeSpent, vLogx_Bookmark, vLogx_Status, vLogx_LastDate, vLogx_SessionId
  Dim vLogx_Eof

  '____ Logx  ___________________________________________

  Sub sReadLogx
    vLogx_No              = oRs("Logx_No")
    vLogx_MembNo          = oRs("Logx_MembNo")
    vLogx_ProgNo          = oRs("Logx_ProgNo")
    vLogx_ModsNo          = oRs("Logx_ModsNo")
    vLogx_Expires         = fNullValue(Trim(fFormatSqlDate(oRs("Logx_Expires"))))
    vLogx_SessionId       = oRs("Logx_SessionId")
    vLogx_Status          = fDefault(oRs("Logx_Status"), 2)
    vLogx_BestDate        = fNullValue(Trim(fFormatSqlDate(oRs("Logx_BestDate"))))
    vLogx_BestScore       = fDefault(oRs("Logx_BestScore"), 0)
    vLogx_NoAttempts      = fDefault(oRs("Logx_NoAttempts"), 0)
    vLogx_TimeSpent       = fDefault(oRs("Logx_TimeSpent"), 0)
    vLogx_Bookmark        = fNullValue(oRs("Logx_Bookmark"))
    vLogx_LastDate        = fNullValue(oRs("Logx_LastDate"))
  End Sub


  Sub sExtractLogx
    vLogx_Expires         = Request("vLogx_Expires")
    vLogx_BestDate        = fNullValue(Trim(Request("vLogx_BestDate")))
    vLogx_BestScore       = fNullValue(Trim(Request("vLogx_BestScore")))
    vLogx_NoAttempts      = fNullValue(Trim(Request("vLogx_NoAttempts")))
    vLogx_TimeSpent       = fNullValue(Trim(Request("vLogx_TimeSpent")))
    vLogx_Bookmark        = fNullValue(Trim(Request("vLogx_Bookmark")))
    vLogx_Status          = Request("vLogx_Status")
    vLogx_LastDate        = fNullValue(Request("vLogx_LastDate"))
  End Sub
  

  Sub sInitLogx
    vLogx_Expires         = Null
    vLogx_BestDate        = Null
    vLogx_BestScore       = Null
    vLogx_NoAttempts      = 0
    vLogx_TimeSpent       = 0
    vLogx_Bookmark        = Null
    vLogx_Status          = 2
    vLogx_LastDate        = Null
    vLogx_SessionId       = Null
  End Sub


  '...get current active log record, ie this will NOT return any expired entries
  Sub spLogxget (vMembNo, vProgNo, vModsNo)
    sOpenCmd
    With oCmd
      .CommandText = "spLogxget"
      .Parameters.Append .CreateParameter("@Logx_MembNo",     adInteger,  adParamInput,    , vMembNo)
      .Parameters.Append .CreateParameter("@Logx_ProgNo",     adInteger,  adParamInput,    , vProgNo)
      .Parameters.Append .CreateParameter("@Logx_ModsNo",     adInteger,  adParamInput,    , vModsNo)
    End With
    Set oRs = oCmd.Execute()
    If oRs.Eof Then 
      sInitLogx
    Else
      sReadLogx
    End If
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...get log record
  Sub spLogxgetByNo (vLogxNo)
    sOpenCmd
    With oCmd
      .CommandText = "spLogxgetByNo"
      .Parameters.Append .CreateParameter("@Logx_MembNo",     adInteger,  adParamInput,    , vLogxNo)
    End With
    Set oRs = oCmd.Execute()
    If oRs.Eof Then 
      sInitLogx
    Else
      sReadLogx
    End If
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...delete log record
  Sub spLogxdeleteByNo (vLogxNo)
    sOpenCmd
    With oCmd
      .CommandText = "spLogxdeleteByNo"
      .Parameters.Append .CreateParameter("@Logx_MembNo",     adInteger,  adParamInput,    , vLogxNo)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub

  '...This posts Expires tracking data for this vMembNo, vProgNo, vModsNo
  '   it assumes that all values are accurate, ie SQL does NOT confirm that the ProgNo exists (or any other value)
  Sub spLogxpost (vMembNo, vProgNo, vModsNo, vExpires, vBestDate, vBestScore, vNoAttempts, vTimeSpent, vBookmark, vStatus, vSessionId, vAcctId, vProgId, vModsId)
    sOpenCmd
    With oCmd
      .CommandText = "spLogxpost"
      .Parameters.Append .CreateParameter("@Logx_MembNo",     adInteger,  adParamInput,    , vMembNo)
      .Parameters.Append .CreateParameter("@Logx_ProgNo",     adInteger,  adParamInput,    , vProgNo)
      .Parameters.Append .CreateParameter("@Logx_ModsNo",     adInteger,  adParamInput,    , vModsNo)
      .Parameters.Append .CreateParameter("@Logx_Expires",    adDBDate,   adParamInput,    , vExpires)
      .Parameters.Append .CreateParameter("@Logx_BestDate",   adDBDate,   adParamInput,    , vBestDate)
      .Parameters.Append .CreateParameter("@Logx_BestScore",  adCurrency, adParamInput,    , vBestScore)
      .Parameters.Append .CreateParameter("@Logx_NoAttempts", adSmallInt, adParamInput,    , vNoAttempts)
      .Parameters.Append .CreateParameter("@Logx_TimeSpent",  adInteger,  adParamInput,    , vTimeSpent)
      .Parameters.Append .CreateParameter("@Logx_Bookmark",   adSmallInt, adParamInput,    , vBookmark)
      .Parameters.Append .CreateParameter("@Logx_Status",     adTinyInt,  adParamInput,    , vStatus)
      .Parameters.Append .CreateParameter("@Logx_SessionId",  adInteger,  adParamInput,    , vSessionId)
      .Parameters.Append .CreateParameter("@Logx_AcctId",     adChar,     adParamInput,   4, vAcctId)
      .Parameters.Append .CreateParameter("@Logx_ProgId",     adChar,     adParamInput,   7, vProgId)
      .Parameters.Append .CreateParameter("@Logx_ModsId",     adChar,     adParamInput,   6, vModsId)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub


  Sub spLogxalterByNo (vLogNo, vStatus, vBestDate, vBestScore, vNoAttempts, vTimeSpent, vBookmark)
    sOpenCmd
    With oCmd
      .CommandText = "spLogxalterByNo"
      .Parameters.Append .CreateParameter("@Logx_No",         adInteger,  adParamInput,    , vLogNo)
      .Parameters.Append .CreateParameter("@Logx_Status",     adTinyInt,  adParamInput,    , vStatus)
      .Parameters.Append .CreateParameter("@Logx_BestDate",   adDBDate,   adParamInput,    , vBestDate)
      .Parameters.Append .CreateParameter("@Logx_BestScore",  adCurrency, adParamInput,    , vBestScore)
      .Parameters.Append .CreateParameter("@Logx_NoAttempts", adSmallInt, adParamInput,    , vNoAttempts)
      .Parameters.Append .CreateParameter("@Logx_TimeSpent",  adInteger,  adParamInput,    , vTimeSpent)
      .Parameters.Append .CreateParameter("@Logx_Bookmark",   adSmallInt, adParamInput,    , vBookmark)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub



  Sub spLogTpost (vStatus, vSessionId, vUrl)
    sOpenCmd
    With oCmd
      .CommandText = "spLogTpost"
      .Parameters.Append .CreateParameter("@LogT_Status",     adVarChar,  adParamInput,   20, vStatus)
      .Parameters.Append .CreateParameter("@LogT_SessionId",  adInteger,  adParamInput,     , vSessionId)
      .Parameters.Append .CreateParameter("@LogT_Url",        adVarChar,  adParamInput, 2000, vUrl)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...Logx Routines 
 
  Function fLogxTimeSpent(vPrev, vMins, vSecs)  
    '...if all values are none then return Null
    If IsNull(vPrev) And vMins = "" And vSecs = "" Then
      fLogxTimeSpent = Null
      Exit Function 
    End If
    '...get any previous TS values - if none vPrev will be null
    vPrev = fPureInt(vPrev)
    vMins = fPureInt(vMins)  
    vSecs = fPureInt(vSecs)  
    '...if we get seconds passed in then these supercede mins (old format)
    If vSecs = 0 Then
      vMins = vMins * 60
    Else
      vMins = 0  
    End If
    fLogxTimeSpent = vPrev + vMins + vSecs      
  End Function 
  

  '...return -1 if not a score between 1 and 100
  Function fScore (vScore)
    fScore = -1    
    If IsNumeric(fOkValue(vScore)) Then
      If vScore >= 0 And vScore <= 1 Then
        fScore = Csng(vScore) * 100
      End If
    End If  
  End Function


  
  '...from Logx.asp - Logx contains ONE record for each Learner/Program/Module/Session
  Function fTrackLogx (vCustId, vAcctId, vMembNo, vProgId, vModsId, vScore, vBookmark, vTimeSpentMins, vTimeSpentSecs, vSessionId, vStatus)
    Dim vMastery, bErr, vMaxAttempts

    fTrackLogx  = ""
    bErr        = False    
    vProgId     = Replace(vProgId, "P0000XX", "") '...there should be no missing ProgIds (Logx will use P0000XX)

    Do While fTrackLogx = ""
      sGetCust vCustId : If vCust_Eof Then fTrackLogx = "No Cust" : bErr = True : Exit Do
      sGetMemb vMembNo : If vMemb_Eof Then fTrackLogx = "No Memb" : bErr = True : Exit Do
      sGetProg vProgId : If vProg_Eof Then fTrackLogx = "No Prog" : bErr = True : Exit Do
      sGetMods vModsId : If vMods_Eof Then fTrackLogx = "No Mods" : bErr = True : Exit Do

      '...get log data up to now (if any) by grabbing any active entries - do not grab expired entries
      spLogxget vMemb_No, vProg_No, vMods_No
  
      '...process tracking parms
      vLogx_TimeSpent  = fLogxTimeSpent (vLogx_TimeSpent, vTimeSpentMins, vTimeSpentSecs)

      '...any bookmarks?
      vLogx_Bookmark = fIf(vLogx_Bookmark > 0 And vBookmark = 0, vLogx_Bookmark, vBookmark)
      
      '...if completed do not modify assessment data
      If vLogx_Status = 3 Then 
        fTrackLogx = "Already Completed"
        Exit Do
      End If


      '...get status from RTE or V5 when deemed "completed" - if nothing passed in then either take from current log or default to 2 (incomplete)
      vLogx_Status = fIf(vStatus > 0, vStatus, fDefault(vLogx_Status, 2))

      '...process valid score (ie other than -1) if not already completed     
      vScore = fScore(vScore)                   '...set to -1 or between 0 and 100 - enters routine as .80
      If vScore = -1 Then
        fTrackLogx = "Content Only"

      Else
        '...get scoring requirements for this Program 
        vMaxAttempts  = 3
        vMastery      = 80
        If vCust_AssessmentAttempts > 0 Then vMaxAttempts = vCust_AssessmentAttempts
        If vCust_AssessmentScore    > 0 Then vMastery     = vCust_AssessmentScore * 100
        If vProg_AssessmentAttempts > 0 Then vMaxAttempts = vProg_AssessmentAttempts
        If vProg_AssessmentScore    > 0 Then vMastery     = vProg_AssessmentScore * 100

        fTrackLogx = "Assessment"

        If vLogx_NoAttempts < vMaxAttempts Then
          vLogx_NoAttempts = vLogx_NoAttempts + 1

          If vScore > vLogx_BestScore Then
            vLogx_BestScore = vScore
            vLogx_BestDate  = Now()
            If vScore >= vMastery Then
              If vStatus = 0 Then vLogx_Status = 3       '...if we did't get a status from previous page (ScoreModule.asp) then set as complete
              If vProg_ResetStatus > 0 Then
                vLogx_Expires = DateAdd("d", vProg_ResetStatus, Now())
                fTrackLogx = "Expired"
              Else
                fTrackLogx = "Completed"
              End If  
            End If
          Else
            fTrackLogx = "Not Best Score"            
          End If
        Else
          fTrackLogx = "Exceeded Attempts"
        End If
      End If
      
    Loop

    '...post valid tracking data
    If Not bErr Then
      spLogxpost vMemb_No, vProg_No, vMods_No, vLogx_Expires, vLogx_BestDate, vLogx_BestScore, vLogx_NoAttempts, vLogx_TimeSpent, vLogx_Bookmark, vLogx_Status, vSessionId, vAcctId, vProgId, vModsId
    End If
    
  End Function    

%>