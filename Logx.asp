<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>

<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<!--#include virtual = "V5/Inc/ProgramStatusRoutines.asp"-->
<!--#include virtual = "V5/Code/ModuleStatusRoutines.asp"-->
<!--#include virtual = "V5/Inc/Fathmail.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->
  
<%
  '...these are all the parameters used in Logs and Logx tracking
  Dim vSource, vScore, vScores, vResults, vBookmark, vLessonLocation, vLessonStatus 
  Dim vTimeSpentTotal, vTimeSpentMins, vTimeSpentSecs, vSessionId, vStatus, vCompleted, vExpired, vTerminate, vPosted

  Dim vSubject, vBody, vSender, vRecipients   
  Dim vBestScore, vNoAttempts, vAssessmentAttempts, vAssessmentScore, vLogsMsg , vLogxMsg
  Dim bScoreOk, bPostOk, bRTE, bCompleted, bTrack

  Const cForReading = 1, cForWriting = 2, cForAppending = 8
  Dim oFs, oF, vFile, vRecord

  vLogsMsg    = ""
  bScoreOk    = False
  bPostOk     = False
  bRTE        = False
  vStatus     = 0         '...get from RTE or V5/ProgramCompleted

'stop  

  '...if from RTE then create session values else set to Null for error tracking
  If fOkValue(Request("vSource")) = "RTE" Then 
    svCustId      = Request("vCustId")
    svMembNo      = Request("vMembNo")
    svCustAcctId  = fCustAcctId(svCustId)    
    bRTE          = True
  ElseIf Not Session("Secure") And Not bRTE Then
    svCustAcctID  = Null
    svMembNo      = Null
    vLogsMsg      = "No Connection"
  End If

' use this to trap fMod postings from closeObjects.asp
' If fOkValue(Request("vSource")) = "CloseObjects" Then Stop

  '...legacy posts
  vProg_Id        = Ucase(Request("vProgId"))
  vMods_Id        = fNullValue (Ucase(Request("vModId")))

  '... if no score set to null else a number    
  vScore          = fOkValue(Request("vScore"))
  If vScore = "" Then
    vScore = NULL
   Else
    vScore = cSng(vScore)
  End If

  vResults            = fNullValue (fUnquote(Request("vResults")))                '...survey result string
  vScores             = Request("vScores")                                        '...v5 scorm - rarely used
  vCompleted          = fDefault(Request("vCompleted"), "n")                      '...from fModules via ProgramCompleted.asp?vCompleted=y - equivalent to RTE vLocationStatus = "completed"
  vTerminate          = fDefault(Request("vTerminate"), "n")                      '...this comes from CloseObjects.asp and allows us to post a terminate to the RTE

  '...RTE posts
  vLessonLocation     = Request("vLessonLocation")                    
  vLessonStatus       = Lcase(Request("vLessonStatus"))               
  vSessionId          = fNullValue(Request("sessionId"))                          '...null or integer (not currently used)
  vPosted             = fFormatSqlDateTime (fDefault(Request("vDate"), Now()))    '...added Jul 2012 - this is the date that the posting occurs, until now it was NOW but this allows posting of old events
  vTimeSpentMins      = fPureInt(Request("vTimeSpent"))                           '...RTE or CloseObjects.asp (v5)
  vTimeSpentSecs      = fMin(fOkInt(fOkInt(Request("vTimeSpentTotalSecs"))/60), 999)      '...RTE new version that relaces all existing TS values (RTE from Jul 2012 onward)     

  If bRTE Then
    vBookmark         = fPureInt(Request("vLessonLocation"))
  Else
    vBookmark         = fPureInt(Request("vPageNo"))
  End If

  '...try and get all the dataset details
  sGetCust svCustId
  sGetMemb svMembNo

  sGetProg vProg_Id
  sGetMods vMods_Id

  '... if any were not on file then do not process this log
  If Not (vCust_Eof Or vMemb_Eof Or vProg_Eof or vMods_Eof) Then
  
    If Session("Secure") Or bRTE Then 
      vLogsMsg = "OK"
      
      '...check score and format it to be a decimal (ie .8) if from RTE
      '   which scores from 0-100 while others post from 0-1, convert percentage to decimal (80 to .80) - it will be converted back to _080 a bit down  
      '   unless it has a wonky stataus or score 

      If IsNumeric(vScore) Then 

        If vScore => 0 Then '...zero is a valid score but not negative numbers
          '...force 1 to be 100 so it doesn't get confused with 1%
          If Not bRTE And vScore = 1 Then 
            vScore = 100 
          End If        
          If vScore => 1 And vScore <= 100 Then 
            vScore = Round(vScore) / 100       
          End If
        End If
        '...consider valid unless the module id is wonky  
        If Len(vMods_Id) = 6 Then
          bScoreOk = True
        Else
          vLogsMsg = "Inv Module Id"          
        End If
      End If
      
      '...did we receive a valid score?
      If bScoreOk Then  
  
        '...fmodules report a score but do not calculate timespent      
        '   trigger a terminate on the RTE
        vTerminate = "y"

        '...get max no of attempts and mastery score (ie .8), else use default values
        vAssessmentAttempts = 3
        vAssessmentScore    = .80      
        If vCust_AssessmentAttempts   > 0 Then vAssessmentAttempts = vCust_AssessmentAttempts
        If vCust_AssessmentScore      > 0 Then vAssessmentScore    = vCust_AssessmentScore
        If Len(vProg_Id) = 7 And vProg_Id <> "P0000XX" Then  '...any by Program?
          If vProg_AssessmentAttempts > 0 Then vAssessmentAttempts = vProg_AssessmentAttempts
          If vProg_AssessmentScore    > 0 Then vAssessmentScore    = vProg_AssessmentScore
        End If   
  
        '...post single score and scores unless we exceded max attempts
        vNoAttempts = fNoAttempts(svMembNo, vMods_Id)
        If vAssessmentAttempts = 99 Or vNoAttempts < vAssessmentAttempts  Then   
  
          '...get best score already on file for this learner, else 0
          vBestScore = fBestTestGrade (svMembNo, vMods_Id) 
          If vBestScore > 0 Then vBestScore = Round(vBestScore) / 100  
  
          '...however if we already have a "passed" score on file, do not record another (added Oct 5, 2010 for Scorm)
          If vBestScore <= vAssessmentScore Then 
  
            vLogs_Item = vMods_Id & "_" & Right("000" & vScore * 100, 3)          
            sLogTestResults2 vPosted                                                                                '...post Scores to LMS (in db_logs.asp) ** uses date
  
            If Not bRTE Then                                                       
              vExpired     = Null
              vCompleted   = "incomplete"
              
              '...completed/passed?          
              If vScore >= vAssessmentScore Then
                vCompleted = "completed"
                '...if recurring assessement send expiry date + # days to reset
                If vProg_ResetStatus > 0 Then
                  vExpired = fFormatSqlDate(DateAdd("d", vProg_ResetStatus, Now()))
                End If
              '...used all attempts?
              Else
                If vNoAttempts + 1 = vAssessmentAttempts Then
                  vCompleted = "failed"
                End If
              End If
  
              '...pass the score and completion status, fRTE will generate the appropriate objective entries        '...post Scores to RTE  
              fRTE vMemb_No, vProg_No, vMods_No, "SetValue", "cmi.core.score.raw", vScore * 100, vExpired, Null, vCompleted        
            End If
    
          vLogsMsg = "Score Ok"
          End If  
          '...post individual scores (not sure if used)
          If Len(vScores) > 0 Then
            vLogs_Item = vMods_Id & "_" & vScores
            sLogAssessmentResults
            vLogsMsg = vLogsMsg & " + ScoreS Ok"
          End If
        Else
          vLogsMsg = vNoAttempts & " attempts exceded"
        End If
      End If
  
  
      '...log completion/timespent/bookmarks/surveys
      If Len(vMods_Id) = 6 Then
  
        '...V5 Surveys (does not post to LMS)
        If Len(vResults) > 0 Then
          vLogs_Item = vProg_Id & "|" & vMods_Id & "_" & vResults
          sLogSurveyResults                                                                                         '...post Surveys to LMS
          vStatus = 3
        End If
  
        '...V5 fModule Completion
        If vCompleted = "y" Then
          sCompletedMod svCustAcctId, svMembNo, vMods_Id                                                            '...post Completion to LMS
          fRTE vMemb_No, vProg_No, vMods_No, "SetValue", "cmi.core.lesson_status", "completed", Null, Null, Null    '...post Completion to RTE
          vStatus = 3
        End If
  
        '...Scorm Bookmarks can be any value (Note: Vubiz modules use page no: 1-999)
        If Len(vLessonLocation) > 0 Then
          fSetLessonLocation vMods_Id, vLessonLocation
        '...see if there's a v5 bookmark from BookmarkObjects.asp
        ElseIf Not bRTE And vBookmark > 0 Then
          fSetBookmark vMods_Id, vBookmark                                                                          '...post BM to LMS
          fRTE vMemb_No, vProg_No, vMods_No, "SetValue", "cmi.core.lesson_location", vBookmark, Null, Null, Null    '...post BM to RTE
        End If    
  
        '...Scorm status is predefined
        If Len(vLessonStatus) > 0 Then
          '...if it's not already completed then set status and completion info
          If Not fLessonCompleted(vMods_Id) Then
            fSetLessonStatus vMods_Id, vLessonStatus
            '...if no score is received and passed or completed - unless score on file, then set as 100%
            If bRTE And (Not bScoreOk) AND (Instr("passed completed", vLessonStatus) > 0) Then
              If vBestScore = 0 Then 
                sCompletedMod svCustAcctId, svMembNo, vMods_Id
              End If
            End If                  
          End If    
          '...pass through location status (complete:3 or incomplete:2)
          vStatus = fIf(Instr("passed completed", vLessonStatus) > 0, 3, 2)
        End If
  

        '...vTimeSpentMins has always been used for LMS and for RTE until Jul 2012
        '...typically there is only one TS at end of a module, so send an initialize in case we had an assessment that did a terminate
        If vTimeSpentMins > 0 Then       
          vTimeSpentTotal = fLogTimeSpent(vProg_Id, vMods_Id, vTimeSpentMins)                                     '...post TS to LMS and RTE ** does not need date
          If Not bRTE Then                                                                                        '...post TS to RTE 
             fRTE vMemb_No, vProg_No, vMods_No, "Initialize", Null, Null, Null, Null, Null          
             fRTE vMemb_No, vProg_No, vMods_No, "SetValue", "cmi.core.session_time", fRTEts (vTimeSpentMins), Null, Null, Null 
          End If
        End If

        '...vTimeSpentSecs start in Jul 2012 for the RTE - the difference is that when using this value we replace what's on file
        '   value was converted to minutes at top of page
        If vTimeSpentSecs > 0 Then       
          sLogTimeSpent vProg_Id, vMods_Id, vTimeSpentSecs, vPosted                                               '...post TS to LMS (RTE new) ** uses date
        End If
  
  
        If vTerminate = "y" And Not bRTE Then
          fRTE vMemb_No, vProg_No, vMods_No, "Terminate", Null, Null, Null, Null, Null                            '...post Terminate to RTE          
        End If
  
      End If
  
    End If

  End If

  '...track details an audit file in a folder called ~Audit in the root of V5 (remember to give the folder write access)
  bTrack      = True      '...if True write all transactions to the audit text file
  bTrack      = False     '...do not track transactions to the audit text file

  If bTrack Then
    Set oFs = CreateObject("Scripting.FileSystemObject")
    vFile   = Server.MapPath("~Audit\") & "\" & fFormatSqlDate(Now()) & ".txt"
    Set oF  = oFs.OpenTextFile(vFile, cForAppending, True)
    vRecord = vbCrLf & "Posted:" & Now() & "|AcctId:" & svCustAcctId & "|MembNo:" & svMembNo & "|ProgId:" & vProg_Id & "|ModsId:" & vMods_Id & "|Score:" & vScore & "|BM:" & vBookmark & "|TS:" & vTimeSpentMins & "|TSS:" & vTimeSpentSecs & "|Logx:" & vLogsMsg & "|Logx:" & vLogxMsg & "|URL:" & Request.QueryString.Item
    oF.Write vRecord
  End If

  If Err.Number <> 0 Then
    vSubject      = "Vubiz Logging Error!"
    vBody         = "<br><br>Content Logging error: " & Err.Number & " (" & Err.Description & " - " & Err.Source & ")<br>" & vSql
    vSender       = "Vubiz Platform Alert <info@vubiz.com>"
    vRecipients   = "Peter Bulloch <pbulloch@vubiz.com>"
    Response.Write fFathMail(vSubject, vBody, vSender, vRecipients)  
  End If

  '...if called from vuAssess then this routine is acting as a web service, return "y" to confirm we got score (vuAssess, former as special assessement service is not used)
  If Lcase(Request("vReturn")) = "y" Then Response.Write "y"

  '...(from SurveyResults.asp) if accessible module then close
  If Lcase(Request("vAccess")) = "y" Then Response.Redirect "CloseObjects.asp?vProg_Id=" & vProg_Id & "&vModId=" & vModId & "&vTimeSpent=" & Request("vTimeSpent")  

%>