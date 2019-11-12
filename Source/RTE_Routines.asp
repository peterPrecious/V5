<%
  '...delete an objective (from LMSsync.asp)
  Sub sRTEdelObj(sesObjId)    
  	If IsNumeric(sesObjId) Then	
		  vSql = "DELETE vuGoldSCORM.dbo.SessionObjective WHERE sesObjID = " & sesObjId 
			sOpenDb
		  oDb.Execute(vSql)
			sCloseDb
		End If
  End Sub


  '...uses values from /Inc/RTE.asp
  '...Get Core Session Data for MyContent
  '   Since this is driven by the Catalogue there will always be a record set containing at least the Module ID and Title - but the RTE values will be null if there is no session
  Sub sRTEsessionCore (vMembNo, vProgNo, vModsId, vProgId)
    sOpenCmdRTE
    With oCmdRTE
      .CommandText = "spRTECatl"
      .Parameters.Append .CreateParameter("@MembNo",   adInteger,	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@ProgNo",   adInteger,	adParamInput,   , vProgNo)
      .Parameters.Append .CreateParameter("@ModsId",   adVarChar, adParamInput,  8, vModsId)
    End With
    Set oRsRTE   				= oCmdRTE.Execute()
  	RTE_ProgNo					= vProgNo
  	RTE_ProgId					= vProgId
  	RTE_ModsNo					= oRsRTE("ModsNo")
  	RTE_ModsId					= oRsRTE("ModsId")
		RTE_ModsTitle  			= oRsRTE("ModsTitle")
		RTE_ModsType  			= Ucase(oRsRTE("ModsType"))
		RTE_ModsScript  		= oRsRTE("ModsScript")
		RTE_ModsFullScreen	= oRsRTE("ModsFullScreen")
		RTE_ModsVuCert    	= oRsRTE("ModsVuCert")
    RTE_ModsFeaAcc      = oRsRTE("ModsFeaAcc")
    RTE_ModsFeaAud      = oRsRTE("ModsFeaAud")
    RTE_ModsFeaVid      = oRsRTE("ModsFeaVid")
    RTE_ModsFeaMob      = oRsRTE("ModsFeaMob")
    RTE_ModsFeaHyb      = oRsRTE("ModsFeaHyb")
		RTE_SessionId				= oRsRTE("SessionId")
		RTE_Expires         = oRsRTE("Expires")
    '...If the latest session has an expires date beyond today than it is closed, else open but not yet "re-accessed"
    If IsNull(RTE_Expires) Then 
      RTE_Closed = Null '...does not apply
    Else  
      If RTE_Expires > Now() Then
        RTE_Closed = True
      Else
        RTE_Closed = False
      End If
    End If
    '...if it was closed but is now "open" then prepare to creation of a new session
    If Not RTE_Closed Then
	    RTE_BestScore				= Null
	    RTE_NoAttempts			= 0
	    RTE_AttemptNo 			= 0
	    RTE_LastDate				= Null
	    RTE_TimeSpent				= 0
	    RTE_Bookmark				= Null
	    RTE_Completed				= False
	    RTE_CompletedDate		= Null
	    RTE_Failed				  = False
    Else
		  RTE_BestScore				= oRsRTE("BestScore")	 
			If Not IsNull(RTE_BestScore) Then  '...this was modified Jul 23, 2018 when there was no bestscore but there was a lastscore (typically due to manual editting)
				RTE_BestScore = fPureInt(RTE_BestScore)
			Else
				RTE_BestScore = fPureInt(oRsRTE("LastScore"))
			End If
		  RTE_NoAttempts			= oRsRTE("NoAttempts") : If IsNull(RTE_NoAttempts)            Then RTE_NoAttempts = 0
		  RTE_AttemptNo 			= oRsRTE("AttemptNo")  : If IsNull(RTE_NoAttempts)            Then RTE_NoAttempts = 0
		  RTE_LastDate				= oRsRTE("LastDate")   
  		RTE_TimeSpent				= oRsRTE("TimeSpent")	 : If IsNull(RTE_TimeSpent)             Then RTE_TimeSpent = 0 Else RTE_TimeSpent = Cdbl(RTE_TimeSpent)
		  RTE_Bookmark				= oRsRTE("Bookmark")	 : If Instr("Z X FX", RTE_ModsType) > 0 Then RTE_Bookmark = fPureInt(RTE_Bookmark) Else RTE_Bookmark = Null
		  RTE_Completed				= oRsRTE("Completed")  : If IsNull(RTE_Completed)             Then RTE_Completed = False ELSE RTE_Completed = True
		  RTE_CompletedDate  	= oRsRTE("Completed")  
		  RTE_Failed				  = oRsRTE("Failed")     : If IsNull(RTE_Failed)                Then RTE_Failed = False
    End If
    If NOT RTE_IsLaunchModule Then		
  		If RTE_Completed Then 
  			RTE_Status = "<!--{{-->Completed<!--}}-->"
  		ElseIf IsDate(RTE_LastDate) Then 
  			RTE_Status = "<!--{{-->Reviewed<!--}}-->"
  		Else
  			RTE_Status = "<!--{{-->Not Started<!--}}-->"
  		End If			
  	Else     
  		If RTE_Completed Then 
  			RTE_Status = "<!--{{-->Completed<!--}}-->"
  		ElseIf IsDate(RTE_LastDate) Then 
'  			RTE_Status = "Best Score: " & RTE_BestScore & "%, Attempts: " & RTE_NoAttempts & " of " & fRTE_MaxAttempts
  			RTE_Status = "Best Score: " & RTE_BestScore & "%, Attempts: " & RTE_AttemptNo & " of " & fRTE_MaxAttempts (vProgId)  
   		ElseIf RTE_Failed Then 
   			RTE_Status = "<!--{{-->Failed<!--}}-->"
  		Else
  			RTE_Status = "<!--{{-->No Attempts<!--}}-->"
  		End If			
    End If
    Set oRsRTE 	= Nothing
    Set oCmdRTE = Nothing
    sCloseDbRTE    
  End Sub

  
  '...Get Objectives for ALL Session Data for RTE_ModsStat (allows for multiple expired sessions)
  Sub sRTEsessionObj2 (vMembNo, vProgNo, vModsNo)

    sOpenCmdRTE
    With oCmdRTE
      .CommandText = "spRTEObj"
      .Parameters.Append .CreateParameter("@MembNo",   adInteger,	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@ProgNo",   adInteger,	adParamInput,   , vProgNo)
      .Parameters.Append .CreateParameter("@ModsNo",   adInteger, adParamInput,   , vModsNo)
    End With
    Set oRsRTE = oCmdRTE.Execute()

		RTE_Attempts	= 0
		RTE_Scores		= ""
		RTE_Dates			= ""
		RTE_Expired   = ""

    Do While Not oRsRTE.Eof 
      RTE_Attempts 	= RTE_Attempts + 1
  		RTE_Scores 		= RTE_Scores  & "|" & oRsRTE("Score")
  		RTE_Dates			= RTE_Dates	  & "|" & oRsRTE("Date")
  		RTE_Expired   = RTE_Expired & "|" & fIf(IsNull(oRsRTE("Expires")), " ", oRsRTE("Expires"))

  		oRsRTE.MoveNext  
    Loop

    Set oRsRTE 	= Nothing
    Set oCmdRTE = Nothing
    sCloseDbRTE    
      
  	'...strip leading pipe	
    If RTE_Attempts > 0 Then
  		RTE_Scores 			= Mid(RTE_Scores, 2)
  		RTE_Dates 			= Mid(RTE_Dates, 2)
  		RTE_Expired			= Mid(RTE_Expired, 2)
  	End If    

  End Sub

  
  '...Get Objectives Session Data for Catl (MyContent)
  '   replaced Aug 24 2012 to get history for multiple sessions, in case of expiry (sRTEObj above)
  Sub sRTEsessionObj (vSessionId)

		RTE_Attempts	= 0
		RTE_Scores		= ""
		RTE_Dates			= ""
    If Len(vSessionId) > 0 Then
      vSql = "SELECT * FROM V5_Vubz.dbo.vRTE_Obj WHERE SessionId =  " & vSessionId 
  '  	sDebug				 			
      sOpenDbRTE    
      Set oRsRTE = oDbRTE.Execute(vSql)
      Do While Not oRsRTE.Eof 
        RTE_Attempts 	= RTE_Attempts + 1
  			RTE_Scores 		= RTE_Scores & "|" & oRsRTE("Score")
  			RTE_Dates			= RTE_Dates	 & "|" & oRsRTE("Date")
  			oRsRTE.MoveNext
      Loop
      Set oRsRTE = Nothing
      sCloseDbRTE          
  		'...strip leading pipe	
      If RTE_Attempts > 0 Then
  			RTE_Scores 			= Mid(RTE_Scores, 2)
  			RTE_Dates 			= Mid(RTE_Dates, 2)
  		End If    
    End If
  End Sub


  '... this fRTE_MaxAttempts FUNCTION is NOT extracted from the RTE but determined by the platform (ie whomever calls this)
  '    it is kept here so My Content can use this value in the STATUS
  Function fRTE_MaxAttempts (vProgId)

    vProg_AssessmentAttempts = fProgAttempts(vProgId)

    If vProg_AssessmentAttempts > 0 Then
      fRTE_MaxAttempts = vProg_AssessmentAttempts
    ElseIf vCust_AssessmentAttempts > 0 Then
      fRTE_MaxAttempts = vCust_AssessmentAttempts
    Else
      fRTE_MaxAttempts = 3
    End If
  End Function


  '...added Feb 29, 2016 for My Content to determine if there is a certificate available
  Function fRTEprogramCompletionDate(vMembNo, vProgNo)
    sOpenCmdRTE
    With oCmdRTE
      .CommandText = "spRTEprogramCompletionDate"
      .Parameters.Append .CreateParameter("@MembNo",   adInteger,	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@ProgNo",   adInteger,	adParamInput,   , vProgNo)
    End With
    Set oRsRTE   				        = oCmdRTE.Execute()
    If oRsRTE.EOF Then
      fRTEprogramCompletionDate = null
    Else
    	fRTEprogramCompletionDate	= oRsRTE("programCompletionDate")
    End If
    Set oRsRTE = Nothing
    Set oCmdRTE = Nothing
  End Function

  '...added Apr 21, 2016 for My Content to BETTER determine if there is a certificate available
  Function fRTEmoduleCompletionDate(vMembNo, vProgNo, vModsNo)
    sOpenCmdRTE
    With oCmdRTE
      .CommandText = "spRTEmoduleCompletionDate"
      .Parameters.Append .CreateParameter("@MembNo",   adInteger,	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@ProgNo",   adInteger,	adParamInput,   , vProgNo)
      .Parameters.Append .CreateParameter("@ModsNo",   adInteger,	adParamInput,   , vModsNo)
    End With
    Set oRsRTE   				        = oCmdRTE.Execute()
    If oRsRTE.EOF Then
      fRTEmoduleCompletionDate = null
    Else
    	fRTEmoduleCompletionDate	= oRsRTE("moduleCompletionDate")
    End If
    Set oRsRTE = Nothing
    Set oCmdRTE = Nothing
  End Function



%>
