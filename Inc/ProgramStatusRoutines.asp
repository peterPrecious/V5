<%
  '...functions/subs for Program Status report

  '...find total Time Spent in a Module (by a given user)
  Function fModTimeSpent(vMembNo, vModId)
    Dim vSql, oRs2
    fModTimeSpent = 0
    sOpenDb2
    vSql = "SELECT RIGHT(Logs_Item, 5) AS TimeSpent FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (RIGHT(LEFT(Logs_Item, 14), 6) = '" & vModId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fModTimeSpent = Clng(oRs2("TimeSpent"))
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...find total Time Spent in a Program (by a given user)
  Function fProgTimeSpent(vMembNo, vProgId)
    Dim vSql, oRs2
    sOpenDb2
    vSql = "SELECT Logs_Item FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (LEFT(Logs_Item, 7) = '" & vProgId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    fProgTimeSpent = 0
    Do While Not oRs2.Eof
      fProgTimeSpent = fProgTimeSpent + Clng(Right(oRs2("Logs_Item"), 6))
      oRs2.MoveNext	        
    Loop
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...find total Number of Attempts at a Test (by a given user)
  Function fNoTestAttempts(vMembNo, vModId)
    Dim vSql, oRs2
    fNoTestAttempts = 0
    sOpenDb2
    vSql = "SELECT COUNT(*) AS NoTestAttempts FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (LEFT(Logs_Item, 6) = '" & vModId & "') ORDER BY NoTestAttempts DESC"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fNoTestAttempts = oRs2("NoTestAttempts")
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...find Best Score Test Grade by a given user) ...used in logx.asp
  '   added date check for recurring scores 
  '   do we need to watch for 1234XX? - see equivalent fBestScore in ModuleStatusRoutines.asp
  Function fBestTestGrade (vMembNo, vModId)
    Dim vSql, oRs2
    fBestTestGrade = 0
    sOpenDb2
    vSql = " SELECT  RIGHT(Logs_Item, 3) AS BestTestGrade"_
         & "   FROM  Logs WITH (NOLOCK)"_ 
         & " WHERE   (Logs_MembNo = " & vMembNo & ")"_ 
         & "   AND   (Logs_Type = 'T') "_ 
         & "   AND   (LEFT(Logs_Item, 6) = '" & vModId & "')"_
         &     fIf(vProg_ResetStatus = 0, "", " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')") _
         &     fIf(vCust_ResetStatus = 0, "", " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')") _
         & " ORDER BY BestTestGrade DESC"
'   Response.Write  vSql
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fBestTestGrade = Clng(oRs2("BestTestGrade"))
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...find total Number of Attempts at an Exam (by a given user)
  Function fNoExamAttempts (vMembNo, vModId)
    Dim vSql, oRs2
    fNoExamAttempts = 0
    sOpenDb2
    vSql = "SELECT RIGHT(LEFT(Logs_Item, 8), 1) AS NoExamAttempts FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'E') AND (LEFT(Logs_Item, 6) = '" & vModId & "') ORDER BY NoExamAttempts DESC"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fNoExamAttempts = Clng(oRs2("NoExamAttempts"))
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...find Best Exam Grade (by a given user)
  Function fBestExamGrade (vMembNo, vModId)
    Dim vSql, oRs2
    fBestExamGrade = 0
    sOpenDb2
    vSql = "SELECT RIGHT(Logs_Item, 2) AS BestExamGrade FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (LEFT(Logs_Item, 6) = '" & vModId & "') ORDER BY BestExamGrade DESC"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fBestExamGrade = Clng(oRs2("BestExamGrade"))
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...post date that user completed this module and generate a score of 100%
  Sub sCompletedMod (vAcctId, vMembNo, vModId)
    Dim vSql, oRs2
    sOpenDb2

    '...insert or update logs for the completion flag
    vSql = "SELECT * FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vModId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then
      vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_MembNo, Logs_Posted, Logs_Item) VALUES ('" & vAcctId & "', 'S', " & vMembNo & ", '" & fFormatSqlDate(Now()) & "', '" & vModId & "')"
    Else
      vSql = "UPDATE Logs SET Logs_Posted = '" & fFormatSqlDate(Now()) & "' WHERE Logs_No = " & oRs2("Logs_No")
    End If
    oDb2.Execute vSql

    vSql = "SELECT * FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Logs_Item = '" & vModId & "_100')"
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then
      vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_MembNo, Logs_Posted, Logs_Item) VALUES ('" & vAcctId & "', 'T', " & vMembNo & ", '" & fFormatSqlDate(Now()) & "', '" & vModId & "_100')"
    Else
      vSql = "UPDATE Logs SET Logs_Posted = '" & fFormatSqlDate(Now()) & "' WHERE Logs_No = " & oRs2("Logs_No")
    End If
    oDb2.Execute vSql

    sCloseDb2
    Set oRs2 = Nothing
  End Sub
  

  '...post date that user completed this program - note above completed modules generate a score as well, but not the Program
  Sub xx_sCompletedProg(vAcctId, vMembNo, vProgId) ' - no longer used
    Dim vSql, oRs2
    sOpenDb2
    vSql = "SELECT * FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vProgId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then
      vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_MembNo, Logs_Posted, Logs_Item) VALUES (" & vAcctId & ", 'S', " & vMembNo & ", '" & fFormatSqlDate(Now()) & "', '" & vProgId & "')"
      oDb2.Execute vSql
    End If
    sCloseDb2
    Set oRs2 = Nothing
  End Sub  


  '...unpost date that user completed this module
  Sub xx_sUnCompletedProg(vAcctId, vMembNo, vProgId) ' - no longer used
    Dim vSql, oRs2
    sOpenDb2
    vSql = "SELECT * FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vProgId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then
      vSql = "DELETE Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vProgId & "')"
      oDb2.Execute vSql
    End If
    sCloseDb2
    Set oRs2 = Nothing
  End Sub  
  
  
  '...find if module is completed (note: originally by a given user, but now if a test exists)
  Function fIsCompletedMod(vMembNo, vModId)
    Dim vSql, oRs2
    fIsCompletedMod = True
    sOpenDb2
    vSql = "SELECT * FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then fIsCompletedMod = False
    sCloseDb2
    Set oRs2 = Nothing
  End Function
  
  
    '...find module lesson status for Scorm (Logs_Type = "L")
  Function fLessonStatus(vMembNo, vModId)
    Dim vSql, oRs2
    fLessonStatus = "Incomplete"
    sOpenDb2
    vSql = "SELECT Logs_Item FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'L') AND (Left(Logs_Item, 6) = '" & vModId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fLessonStatus = Mid(oRs2("Logs_Item"), 8)
    sCloseDb2
    Set oRs2 = Nothing
  End Function
  
  
  '...find if assessment on file
  Function fAssessment (vMembNo, vModId)
    Dim vSql, oRs2
    fAssessment = ""
    sOpenDb2
    vSql = "SELECT * FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    Do While Not oRs2.Eof
      If fAssessment <> "" Then fAssessment = fAssessment & "<br>"
      fAssessment = fAssessment & vModId & " - " & fFormatSqlDate(oRs2("Logs_Posted")) & " - " & FormatNumber(Right (oRs2("Logs_Item"), 3), 0)
      oRs2.MoveNext      
    Loop
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...find if program is completed (by a given user) ' - no longer used
  Function xx_fIsCompletedProg(vMembNo, vProgId)
    Dim vSql, oRs2
    sOpenDb2
    vSql = "SELECT Logs_Posted FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vProgId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then 
      fIsCompletedProg = "N"
    Else
      fIsCompletedProg = "Y"
      vLogs_Posted = oRs2("Logs_Posted")  
    End If
    sCloseDb2
    Set oRs2 = Nothing
  End Function
 
 
  '...find if module has been accessed (by a given user)
  Function fIsAccessedMod(vMembNo, vModId)
    Dim vSql, oRs2
    fIsAccessedMod = True
    sOpenDb2
    vSql = "SELECT RIGHT(Logs_Item, 5) AS TimeSpent FROM Logs WITH (NOLOCK) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (RIGHT(LEFT(Logs_Item, 14), 6) = '" & vModId & "')"
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then fIsAccessedMod = False
    sCloseDb2
    Set oRs2 = Nothing
  End Function


  '...log timespent in the module - if no Program use P0000XX - minimum 1 min
  Function fLogTimeSpent(vProgId, vModId, vTimeSpentMins)
    Dim vTimeSpent '...put into different variable so we don't return total value as original is needed for Logx
    vTimeSpent = vTimeSpentMins
    vProgId    = Ucase(fOkValue(vProgId))
    If Len(vProgId) <> 7 Then 
      vProgId = "P0000XX"
    ElseIf Left(vProgId, 1) <> "P" Then 
      vProgId = "P0000XX"
    End If
    vModId    = fOkValue(vModId)
    If Len(vModId) = 6 Then 
      vTimeSpent = fMax(Clng(vTimeSpent), 1) '...default to 1 minute  
      sOpenDb
      '...see if course on file
      vSql = "SELECT Logs_Item FROM Logs WITH (NOLOCK) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'P' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 14) = '" & vProgId & "|" & vModId & "'"
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then     
        '...add to existing timespent
        vTimeSpent = vTimeSpent + Clng(Right(oRs("Logs_Item"), 6))
        vLogs_Item = vProgId & "|" & vModId & "_" & Right("000000" & vTimeSpent, 6) 
        '...update existing timespent
        vSql = " UPDATE Logs SET Logs_Item = '" & vLogs_Item & "', Logs_Posted = '" & fFormatSqlDate(Now()) & "'" _
             & " WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'P' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 14) = '" & vProgId & "|" & vModId & "'"
        oDb.Execute(vSql)    
      Else      
        '...add timespent on new course 
        vLogs_Item = vProgId & "|" & vModId & "_" & Right("000000" & vTimeSpent, 6)
        vSql = "INSERT INTO Logs "
        vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
        vSql = vSql & "('" & svCustAcctId & "', 'P', '" & vLogs_Item & "', " & svMembNo & ") "
        oDb.Execute(vSql)
      End If
      Set oRs = Nothing
      sCloseDb
    End If
    fLogTimeSpent = vTimeSpent '...this is the total time spent on file after operation
  End Function


  '...log fx/xx modules timespent from RTE (new Jul 212)
  Sub sLogTimeSpent (vProgId, vModId, vTSmins, vPosted)
    sOpenDb
    '...delete if on file
    vSql = "DELETE Logs WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'P' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 14) = '" & vProgId & "|" & vModId & "'"
    oDb.Execute(vSql)
    '...add in new TS record
    vLogs_Item = vProgId & "|" & vModId & "_" & Right("000000" & vTSmins, 6)
    vSql = "INSERT INTO Logs "
    vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo, Logs_Posted) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'P', '" & vLogs_Item & "', " & svMembNo & ", '" & vPosted & "') "
    oDb.Execute(vSql)
    Set oRs = Nothing
    sCloseDb
  End Sub


  '...set a V5 Bookmark for a Module.  Used in BookmarkObjects.asp / ScoreModule.asp
  Function fSetBookmark(vModId, vBookmark)
    If Len(vModId) = 6 And vBookmark > 0 Then  
      sOpenDb
      vLogs_Item = vModId & "_" & Right("0000" & vBookmark, 3)  
      '...see if already on file
      vSql = "SELECT Logs_No FROM Logs WITH (NOLOCK) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'B' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 6) = '" & Left(vLogs_Item, 6) & "'"
      Set oRs = oDb.Execute(vSql)
      If Not oRs.EOF Then      
        '...add to existing bookmark
        vSql = "UPDATE Logs SET Logs_Item = '" & vLogs_Item & "', Logs_Posted = '" & fFormatSqlDate(Now) & "' WHERE Logs_No= " & oRs("Logs_No")
        oDb.Execute (vSql)    
      Else      
        vSql = "INSERT INTO Logs "
        vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
        vSql = vSql & "('" & svCustAcctId & "', 'B', '" & vLogs_Item & "', " & svMembNo & ") "
        oDb.Execute (vSql)  
      End If
      sCloseDb
      Set oRs = Nothing  
    End If
  End Function


  '...see if LessonStatus is "completed"
  Function fLessonCompleted(vModId)
    fLessonCompleted = false
    If Len(vModId) = 6 Then  
      sOpenDb
      vSql = "SELECT Logs_No FROM Logs WITH (NOLOCK) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'L' AND Logs_MembNo = " & svMembNo & " AND Logs_Item = '" & vModId & "_completed'"
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then fLessonCompleted = True
      sCloseDb
      Set oRs = Nothing  
    End If
  End Function


  '...set the LessonStatus for a Scorm Module.  Used in LaunchObjects.asp and ScoreModule.asp
  '   should be one of: passed completed failed incomplete browsed not attempted
  Function fSetLessonStatus(vModId, vLessonStatus)
    If Len(vModId) = 6 And Len(vLessonStatus) > 0 Then  
      sOpenDb
      vLogs_Item = vModId & "_" & vLessonStatus
      '...see if already on file
      vSql = "SELECT Logs_No FROM Logs WITH (NOLOCK) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'L' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 6) = '" & Left(vLogs_Item, 6) & "'"
      Set oRs = oDb.Execute(vSql)
      If Not oRs.EOF Then      
        vSql = "UPDATE Logs SET Logs_Item = '" & vLogs_Item & "', Logs_Posted = '" & fFormatSqlDate(Now()) & "' WHERE Logs_No= " & oRs("Logs_No")
        oDb.Execute (vSql)    
      Else      
        vSql = "INSERT INTO Logs "
        vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
        vSql = vSql & "('" & svCustAcctId & "', 'L', '" & vLogs_Item & "', " & svMembNo & ") "
        oDb.Execute (vSql)  
      End If
      sCloseDb
      Set oRs = Nothing  
    End If
  End Function


  '...set a LessonLocation for a Scorm Module.  Used in ScoreModule.asp
  Function fSetLessonLocation(vModId, vBookmark)
    If Len(vModId) = 6 And Len(vBookmark) > 0 Then  
      sOpenDb
      vLogs_Item = vModId & "_" & vBookmark
      '...see if already on file
      vSql = "SELECT Logs_No FROM Logs WITH (NOLOCK) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'B' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 6) = '" & Left(vLogs_Item, 6) & "'"
      Set oRs = oDb.Execute(vSql)
      If Not oRs.EOF Then      
        '...add to existing bookmark
        vSql = "UPDATE Logs SET Logs_Item = '" & vLogs_Item & "', Logs_Posted = '" & fFormatSqlDate(Now) & "' WHERE Logs_No= " & oRs("Logs_No")
        oDb.Execute (vSql)    
      Else      
        vSql = "INSERT INTO Logs "
        vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
        vSql = vSql & "('" & svCustAcctId & "', 'B', '" & vLogs_Item & "', " & svMembNo & ") "
        oDb.Execute (vSql)  
      End If
      sCloseDb
      Set oRs = Nothing  
    End If
  End Function

 
%>
