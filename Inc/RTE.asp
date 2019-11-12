<%
	'...use these for current session
	Dim RTE_Expires, RTE_Closed, RTE_ProgNo, RTE_ProgId, RTE_SessionId, RTE_BestScore, RTE_LastDate, RTE_Completed, RTE_CompletedDate, RTE_Failed, RTE_Bookmark, RTE_TimeSpent, RTE_Status, RTE_IsLaunchModule
	Dim RTE_ModsNo, RTE_ModsId, RTE_ModsTitle, RTE_ModsType, RTE_ModsScript, RTE_ModsFullScreen, RTE_ModsFeaAcc, RTE_ModsFeaAud, RTE_ModsFeaVid, RTE_ModsFeaMob, RTE_ModsFeaHyb, RTE_ModsVuCert   

  '...from the objectives table and used in RTE_Routines for RTE_ModsStat
	Dim RTE_Scores, RTE_Dates, RTE_Expired

  '... RTE_AttemptNo is the current number of attempts taken, can be zero (if just reset) or a max of the amount - added Aug 1, 2012
  '... RTE_NoAttempts is extracted from core the same was as RTE_Attemps is calculated - added Aug 1, 2012
  '... RTE_Attempts is the same as RTE_NoAttemps (count of Objectives scores) but return when rendering history
  '... RTE_AttemptsLeft is max attempts from platform less RTE_AttemptNo - used to render links on My Content - added Aug 1, 2012 (NOTE: computed not db driven)
  '... RTE_Max Attempts is used in the RTE_ModsStat
  Dim RTE_NoAttempts, RTE_Attempts, RTE_AttemptNo, RTE_AttemptsLeft
  '--, RTE_MaxAttempts

  '... this stores the page that calls the RTE routines and is passed/stored in suspend_data
  Session("RTE_Caller") = svPage





  Function fRTE (vMembNo, vProgNo, vModsNo, vOp, vParm, vValue, vExpires, vPostDate, vCompleted)

    Dim vGuidId, vInitId, vGuid, bInit, vUrl, vOpenDate, vSend, vResp, vObjective

		'.../web01/ must be setup in local hosts file
'   vUrl    		= "//localhost/Gold/vuSCORM/service.asmx/"    
'   vUrl    		= "//staging.vubiz.com/Gold/vuSCORM/service.asmx/"    
'   vUrl    		= "//web01/Gold/vuSCORM/service.asmx/"    

    '...get FULL URL for the Web Service
    If svSSL Then 
      Select Case svServer
        Case "cloudweb.vubiz.com"       : vUrl = "https://cloudweb.vubiz.com/Gold/vuSCORM/service.asmx/" 
        Case "stagingweb.vubiz.com"     : vUrl = "https://stagingWeb.vubiz.com/Gold/vuSCORM/service.asmx/" 
        Case "localhost"                : vUrl = "https://stagingWeb.vubiz.com/Gold/vuSCORM/service.asmx/" 
        Case Else                       : vUrl = "https://vubiz.com/Gold/vuSCORM/service.asmx/"     
      End Select  
    Else    
      Select Case svServer
        Case "cloudweb.vubiz.com"       : vUrl = "http://cloudweb.vubiz.com/Gold/vuSCORM/service.asmx/" 
        Case "stagingweb.vubiz.com"     : vUrl = "http://stagingWeb.vubiz.com/Gold/vuSCORM/service.asmx/" 
        Case "localhost"                : vUrl = "http://stagingWeb.vubiz.com/Gold/vuSCORM/service.asmx/" 
        Case Else                       : vUrl = "http://vubiz.com/Gold/vuSCORM/service.asmx/"     
      End Select
    End If


    fRTE = ""

    '...get guid for quick fixes (LearnerReportCard_ws.asp) then exit
    If vOp = "GetSessionGuid" Then
      vSend    = vUrl & "GetSessionGUID?MemberID=" & vMembNo & "&ModuleID=" & vModsNo & "&ProgramID=" & vProgNo & "&SCOID=&SCOTitle="
      fRTE     = fRTEws(vSend) 
      Exit Function
    End If

    '...for each operation requested, get the Session GUID for this Memb/Prog/Mods 
    '...format: Session("guid_123_456") where 123=progno and 456=modsno, minimum
    '...and whether that session has been initialized or not (Session("init_123_4567")

    vGuidId			= "guid" & "_" & vProgNo & "_" & vModsNo
    vInitId			= "init" & "_" & vProgNo & "_" & vModsNo

    vGuid				= fOkValue(Session(vGuidId))					'...vGuid will only exists if initialized in an earlier call
    bInit				= fDefault(Session(vInitId), False)		'...bGuid is false unless initialized

		'...do we have the guid in a session variable or do we need to get it from the RTE
		If Len(vGuid) = 0 Then
      vSend    = vUrl & "GetSessionGUID?MemberID=" & vMembNo & "&ModuleID=" & vModsNo & "&ProgramID=" & vProgNo & "&SCOID=&SCOTitle="
      vResp    = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
      If vResp = "ERR" Then Exit Function
      vGuid  	 = vResp : Session(vGuidId) = vGuid
      bInit		 = False : Session(vInitId) = bInit

      '...store the V5 source into the session data to differentiate from SCORM (not critical if fails)
      vSend    = vUrl & "SetValue?SessionGUID=" & vGuid & "&Parameter=vubiz.scorm_version&Value=V5"
      vResp    = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
	  End If  


		'...this comes from LaunchObjects.asp (another initialize can be from setting scores below
    If vOp = "Initialize" And bInit = False Then 
  		If Not fInitialize (vMembNo, vProgNo, vModsNo, vGuid, bInit, vUrl) Then Exit Function
      bInit	= True : Session(vInitId) = bInit
    End If


    If vOp = "SetValue" Then

      '...if date is passed in and it's less than session date then set to historical date (format Jan 1, 2005)
      '   this date is normally in real time but it's the date the session was created
      If IsDate(vPostDate) Then
        vSend      = vUrl & "GetValue?SessionGUID=" & vGuid & "&Parameter=vubiz.open_date"
        vResp      = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
        If vResp   = "ERR" Then Exit Function

        vOpenDate  = fIf(IsDate(vResp), fFormatSqlDate(vResp), fFormatSqlDate(Now))      
        '...only set the open_date if it's older than what we currently have
        If DateDiff("d", vPostDate, vOpenDate) > 0 Then      
          vSend    = vUrl & "SetValue?SessionGUID=" & vGuid & "&Parameter=vubiz.open_date&Value=" & vPostDate
          vResp    = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
          If vResp = "ERR" Then Exit Function
        End If
      End If


			'...if we are posting a score then assume a new module has been initialized (ie launched frorm another module) and terminated 
			'   as well as handle the objectives
			If vParm = "cmi.core.score.raw" Then

				'...the assessment version of this module has been launched from a module link then initialize if not already
		    If bInit = False Then 
					If Not fInitialize (vMembNo, vProgNo, vModsNo, vGuid, bInit, vUrl) Then Exit Function
		      bInit	= True : Session(vInitId) = bInit
			  End If

  			'...get the objective count (ie the no of scores already posted)
				vSend = vUrl & "GetValue?SessionGUID=" & vGuid & "&Parameter=cmi.objectives._count"
	      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
	      If vResp = "ERR" Then Exit Function
	      vObjective = vResp

				'...post the score in the objective
	      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=cmi.objectives." & vObjective & ".score.raw&Value=" & vValue 
	      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
	      If vResp = "ERR" Then Exit Function

				'...post completion status in the objective
	      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=cmi.objectives." & vObjective & ".completion_status&Value=" & vCompleted
	      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
	      If vResp = "ERR" Then Exit Function

				'...post date in objective if historical
				If DateDiff("d", vPostDate, Now()) > 0 Then   
		      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=cmi.objectives." & vObjective & ".timestamp&Value=" & vPostDate
		      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
	      End If
	      If vResp = "ERR" Then Exit Function

				'...post the score in the core
	      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=" & vParm & "&Value=" & vValue 
	      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
	      If vResp = "ERR" Then Exit Function

				'...post the status in the core
	      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=cmi.core.lesson_status&Value=" & vCompleted
	      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
	      If vResp = "ERR" Then Exit Function

				'...post expired status in the core
	      If IsDate(vExpires) Then
		      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=vubiz.expiry_date&Value=" & vExpires
		      vResp = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
		      If vResp = "ERR" Then Exit Function
		    End If

			'...if we are NOT posting a score than post normally
			Else

	      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=" & vParm & "&Value=" & vValue 
	      vResp= fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
	      If vResp = "ERR" Then Exit Function
			End If


    ElseIf vOp = "GetValue" Then

      vSend = vUrl & vOp & "?SessionGUID=" & vGuid & "&Parameter=" & vParm
      vResp = fRTEws(vSend)      
      If vResp = "ERR" Then Exit Function

    ElseIf vOp = "Terminate" Then

	    If Not fTerminate (vMembNo, vProgNo, vModsNo, vGuid, vUrl) Then Exit Function
	    '...turn off the init value
			bInit	= False : Session(vInitId) = bInit

    End If

  End Function


	Function fInitialize(vMembNo, vProgNo, vModsNo, vGuid, bInit, vUrl)
		Dim vSend, vResp
		fInitialize = True    

		'...initialize session
    vSend 		= vUrl & "Initialize?SessionGUID=" & vGuid
    vResp 		= fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
    If vResp  = "ERR" Then fInitialize = False : Exit Function

    '...save the calling routine for debugging
    vSend     = vUrl & "SetValue?SessionGUID=" & vGuid & "&Parameter=cmi.suspend_data&Value=" & Session("RTE_Caller")
    vResp     = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
    If vResp 	= "ERR" Then fInitialize = False : Exit Function 

    '...if status unknown the set to incomplete (don't kill if fails)
    vSend     = vUrl & "GetValue?SessionGUID=" & vGuid & "&Parameter=cmi.core.lesson_status"
    vResp     = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
    If vResp 	= "ERR" Then fInitialize = False : Exit Function 

		If vResp <> "not attempted" Then Exit Function
    vSend   	= vUrl & "SetValue?SessionGUID=" & vGuid & "&Parameter=cmi.core.lesson_status&Value=incomplete"
    vResp   	= fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp      
    If vResp  = "ERR" Then fInitialize = False : Exit Function

	End Function 


	Function fTerminate(vMembNo, vProgNo, vModsNo, vGuid, vUrl)
		Dim vSend, vResp
   	fTerminate = True
    vSend 		 = vUrl & "Terminate?SessionGUID=" & vGuid
    vResp 		 = fRTEws(vSend) : sTrack vMembNo, vProgNo, vModsNo, vSend, vResp
    If vResp 	 = "ERR" Then fTerminate = False
	End Function


	'...create a CSV record for each call to the RTE
  Sub sTrack (vMembNo, vProgNo, vModsNo, vSend, vResp)
    Dim bTrack : bTrack = False '...disable if CSV tracking no longer needed - ensure there's a folder called \V5\~RTE\ if tracking
    If bTrack Then
	    Const cForReading = 1, cForWriting = 2, cForAppending = 8
	    Dim oFs, oF, vFile, vRec, vUrlLen
      Set oFs = CreateObject("Scripting.FileSystemObject")
      vFile   = Server.MapPath("\V5\~RTE\") & "\" & fFormatSqlDate(Now()) & ".csv"
      Set oF  = oFs.OpenTextFile(vFile, cForAppending, True)
    	vUrlLen			= Instr(vSend, ".asmx") + 6 '...strip off the base URL
      oF.Write vbCrLf & Now() & "," & vMembNo & "," & vProgNo & "," & vModsNo & "," & Mid(vSend, vUrlLen) & "," & vResp 
    End If	
	End Sub


	'...this converts minutes into HH:MM:SS format
	'	i =  36 : Response.Write "<br>" & i & " secs = " & fRTEts (i)
	'	i =  60 : Response.Write "<br>" & i & " secs = " & fRTEts (i)
	'	i =   0 : Response.Write "<br>" & i & " secs = " & fRTEts (i)
	'	i =   1 : Response.Write "<br>" & i & " secs = " & fRTEts (i)
	'	i = 600 : Response.Write "<br>" & i & " secs = " & fRTEts (i)
	Function fRTEts (vMM)
		fRTEts = Right("00" & Cint(vMM\60), 2) & ":" & Right("00" & vMM Mod 60, 2) & ":00"
	End Function


  '...launch the web service (vURL must be a full url, ie http://.... or https://....)
  Function fRTEws(vUrl)    
    fRTEws = "ERR"
    Dim oWs
    Set oWs = Server.CreateObject("MSXML2.XMLHTTP")
    oWs.Open "GET", vUrl, false
'   oWs.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
    oWs.Send vUrl
    fRTEws = oWs.ResponseXML.Text
    Set oWs = Nothing                
  End Function

%>
