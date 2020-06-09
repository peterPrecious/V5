<%
  '...build the status link for My Content and My World
  Function fModStatusLink (vMembNo, vProgId, vModId)
    Dim vTitle  
    sGetProg (vProg_Id)  '...get prog info as it was probably NOT retrieved in MyWorld
    vTitle = Server.HtmlEncode(fPhraH(001357)) & ": " & fModStatus (vMembNo, vModId) '...Status
    vTitle = "<span class='green'>" & vTitle & "</span>"  
    fModStatusLink = "<a href=""javascript:vuwindow('StatusModules.asp?vClose=Y&vProgId=" & vProgId & "&vModId=" & vMods_Id & "',500,450,50,50,'no','no','yes')"">" & vTitle & "</a>"
  End Function
  

  '...determine module status 
  Function fModStatus (vMembNo, vModId)
    If fCompleted (vMembNo, vModId) <> "" Then
      fModStatus  = Server.HtmlEncode(fPhraH(000107)) '...Completed
    ElseIf fTimeSpent(vMembNo, vModId) > 0 Then
      fModStatus  = Server.HtmlEncode(fPhraH(000229)) '...Reviewed
    Else
      fModStatus  = Server.HtmlEncode(fPhraH(000192)) '...Not Started
    End If
  End Function


  '...return date module was either flagged as completed by V5 or status = "complete" by Scorm or assessment >= 80 over past x days
  Function fCompleted (vMembNo, vModId)

   'If vMembNo = 2339117 And vModId = "4768EN" Then Stop
   'If vModId = "40374EN" Then Stop

    '...check v5 for completion status
    sOpenDb3
    vSql = "SELECT Logs_Posted FROM Logs "

    If Ucase(Right(vModId, 2)) = "XX" Then
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Left(Logs_Item, 4) = '" & Left(vModId, 4) & "')"
    Else
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vModId & "')"
    End If

    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If

    Set oRs3 = oDb3.Execute(vSql)
    If oRs3.Eof Then 
      fCompleted = ""
    Else
      fCompleted = oRs3("Logs_Posted")
    End If
    sCloseDb3
    Set oRs3 = Nothing

    '...check Scorm status for complete
    If fCompleted = "" Then
      sOpenDb3
      vSql = "SELECT Logs_Posted FROM Logs "

      If Ucase(Right(vModId, 2)) = "XX" Then
        vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'L') AND (Logs_Item LIKE '" & Left(vModId, 4) & "___completed')"
      Else  
        vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'L') AND (Logs_Item = '" & vModId & "_completed')"
      End If

      '...added to ensure status reflects only activity within last x days as defined in the customer table
      If vProg_ResetStatus > 0 Then     
        vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
      ElseIf vCust_ResetStatus > 0 Then     
        vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
      End If
      Set oRs3 = oDb3.Execute(vSql)
      If oRs3.Eof Then 
        fCompleted = ""
      Else
        fCompleted = oRs3("Logs_Posted")
      End If
      sCloseDb3
      Set oRs3 = Nothing
    End If
    
    '...else see if best score is a "pass" 
    If fCompleted = "" Then
      If fBestScore(vMembNo, vModId)/100 >= fPassingScore() Then
        fCompleted = fLastScore(vMembNo, vModId)
      End If      
    End If    
   
  End Function


  '...return date survey was taken
  Function fSurveyCompleted (vMembNo, vModId)
'   If Len(vModId) = 6 Then vModId = "undefined|" & vModId
    sOpenDb3
    vSql = "SELECT Logs_Posted FROM Logs "
    vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'U') AND  (RIGHT(LEFT(Logs_Item, 14), 6) = '" & vModId & "')"
    Set oRs3 = oDb3.Execute(vSql)
    If oRs3.Eof Then 
      fSurveyCompleted = ""
    Else
      fSurveyCompleted = oRs3("Logs_Posted")
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Function


  '...get best score allowing for reset (modified Apr 19, 2016 to handle big mods)
  Function fBestScore (vMembNo, vModId)
    vSql = "SELECT MAX(CAST(Right(Logs.Logs_Item, 3) AS FLOAT)) AS Logs_Grade FROM Logs "
    If Ucase(Right(vModId, 2)) = "XX" Then
'     vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 4) = '" & Left(vModId, 4) & "')"
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, LEN(Logs_Item) - 6) = '" & Left(vModId, LEN(vModId) - 2) & "')"
    Else
'     vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "')"
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, LEN(Logs_Item) - 4) = '" & vModId & "')"
    End If
    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If
'   Response.Write  vSql
    sOpenDb3
    Set oRs3 = oDb3.Execute(vSql)
    If oRs3.Eof Then 
      fBestScore = -1
    ElseIf IsNull(oRs3("Logs_Grade")) Then
      fBestScore = -1
    Else
      fBestScore = Cint(oRs3("Logs_Grade"))
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Function


  '...get first score of module assessment - this is when only 1 attempt is permitted
  Function fFirstScore (vMembNo, vModId)
    vSql = "SELECT TOP 1 CAST(Right(Logs.Logs_Item, 3) AS FLOAT) AS Logs_Grade FROM Logs"
'   vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "')"
    vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, LEN(Logs_Item) - 4) = '" & vModId & "')"


    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If
    sOpenDb3
    Set oRs3 = oDb3.Execute(vSql)
    If oRs3.Eof Then 
      fFirstScore = 0
    ElseIf IsNull(oRs3("Logs_Grade")) Then
      fFirstScore = 0
    Else
      fFirstScore = Cint(oRs3("Logs_Grade"))
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Function


  '...get date of last assessment attempt (twigged Nov 8, 2018 to handle big mods)
  Function fLastScore (vMembNo, vModId)
    vSql = "SELECT MAX(Logs.Logs_Posted) AS Logs_Posted FROM Logs"
    If Ucase(Right(vModId, 2)) = "XX" Then '... ignore language (EN, etc)
 '    vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 4) = '" & Left(vModId, 4) & "')"
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, LEN(Logs_Item) - 6) = '" & Left(vModId, LEN(vModId) - 2) & "')"
    Else
 '    vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "')"
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, LEN(Logs_Item) - 4) = '" & vModId & "')"


    End If
    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If
    sOpenDb3
    Set oRs3 = oDb3.Execute(vSql)
    If oRs3.Eof Then 
      fLastScore = ""
    ElseIf IsNull(oRs3("Logs_Posted")) Then
      fLastScore = ""
    Else
      fLastScore = oRs3("Logs_Posted")
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Function


  '...find total Time Spent in a Module - for MyWorld Status line and VuAssess.com
  Function fTimeSpent(vMembNo, vModId)
    sOpenDb3
'   vSql = "SELECT RIGHT(Logs_Item, 5) AS TimeSpent FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (RIGHT(LEFT(Logs_Item, 14), 6) = '" & vModId & "')"
    vSql = "SELECT RIGHT(Logs_Item, 5) AS TimeSpent FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (CHARINDEX('|" & vModId & "', Logs_Item) > 0)"
    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If
    fTimeSpent = 0
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then
      If Len(oRs3("TimeSpent")) > 0 Then
        fTimeSpent = Clng(oRs3("TimeSpent"))
      End If
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Function


  '...find last Time learner was in a Module - for StatusModules.asp
  Function fLastSpent(vMembNo, vModId)
    sOpenDb3
    vSql = "SELECT Logs_Posted FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (RIGHT(LEFT(Logs_Item, 14), 6) = '" & vModId & "')"
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then
      fLastSpent = fFormatDate(oRs3("Logs_Posted"))
    Else
      fLastSpent = ""
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Function


  '...find if Module is completed (by a given user)
  Function fIsCompleteMod(vMembNo, vModId)
    If fBestScore(vMembNo, vModId) >= 80 Or fCompleted (vMembNo, vModId) <> "" Then
      fIsCompleteMod = True
    Else
      fIsCompleteMod = False
    End If
  End Function


  '...post date that user completed this module (this routine also in ProgramStatusRoutines.asp but called sCompletedMod)
  Sub sCompleteMod(vAcctId, vMembNo, vModId)
    sOpenDb3
    vSql = "SELECT * FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vModId & "')"
    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If
    Set oRs3 = oDb3.Execute(vSql)
    If oRs3.Eof Then
      vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_MembNo, Logs_Posted, Logs_Item) VALUES ('" & vAcctId & "', 'S', " & vMembNo & ", '" & Now() & "', '" & vModId & "')"
      oDb3.Execute vSql
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Sub


  '...unpost date that user completed this module
  Sub sUnCompleteMod(vAcctId, vMembNo, vModId)
    sOpenDb3
    vSql = "SELECT * FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vModId & "')"
    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
'     sDebug
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
'     sDebug
    End If
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then
      vSql = "DELETE Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'S') AND (Logs_Item = '" & vModId & "')"
      oDb3.Execute vSql
    End If
    sCloseDb3
    Set oRs3 = Nothing
  End Sub 


  '...find Max Attempts available to this learner for this assessment
  Function fMaxAttempts(vMembNo, vModId)
    If vProg_AssessmentAttempts > 0 Then
      fMaxAttempts = vProg_AssessmentAttempts
    ElseIf vCust_AssessmentAttempts > 0 Then
      fMaxAttempts = vCust_AssessmentAttempts
    Else
      fMaxAttempts = 3
    End If
  End Function


  '...find total Number of Attempts at assessment  (modified Mar 29, 2016 to handle big mods
  Function fNoAttempts(vMembNo, vModId)
    Dim vSql, oRs3
    fNoAttempts= 0
    sOpenDb3
'   vSql = "SELECT COUNT(*) AS NoAttempts FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (LEFT(Logs_Item, 6) = '" & vModId & "')"
    vSql = "SELECT COUNT(*) AS NoAttempts FROM Logs WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Logs_Item LIKE '" & vModId & "%')"  '...new format (maybe)
    '...added to ensure status reflects only activity within last x days as defined in the customer table
    If vProg_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vProg_ResetStatus * -1, Now)) & "')"
    ElseIf vCust_ResetStatus > 0 Then     
      vSql = vSql & " AND (Logs_Posted > '" & fFormatSqlDate(DateAdd("d", vCust_ResetStatus * -1, Now)) & "')"
    End If
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then fNoAttempts= Cint(oRs3("NoAttempts"))
    sCloseDb3
    Set oRs3 = Nothing
  End Function
  
  
  '...get best score of a traditional exam
  Function fExamStatus (vMembNo, vExamString)
    fExamStatus = Server.HtmlEncode(fPhraH(000360))
    Dim vModId, i, j
    '...typical exam string : "ExamStart.asp?vModID=1177EN&vStart=True&vMinQue=30&vMaxAttempts=4&vBankTLimit=0&vPassGrade=80"
    i = Instr(vExamString, "vModID=")
    If i > 0 Then 
      vModId = Mid(vExamString, i + 7, 6)
      j = fBestScore (vMembNo, vModId)
      If j >= 0 Then fExamStatus = Server.HtmlEncode(fPhraH(000361)) & ":&nbsp;" & j & "%"
      If (svBrowser <> "msie") Then
        fExamStatus = fExamStatus & " | <a onclick='location.reload();' href='#'>" & Server.HtmlEncode(fPhraH(000522)) & "</a>"
      End If
    End If
  End Function  
  

  '...get best score of a module assessment
  Function fAssessmentStatus(vMembNo, vModId)
    Dim i, j, k
    i = fMaxAttempts (vMembNo, vModId)
    j = fBestScore   (vMembNo, vModId)
    k = fNoAttempts  (vMembNo, vModId)   
    If k = 0 Then 
      fAssessmentStatus = Server.HtmlEncode(fPhraH(000360))
    Else
      fAssessmentStatus = Server.HtmlEncode(fPhraH(000361)) & ":&nbsp;" & j & "%"
      fAssessmentStatus = fAssessmentStatus & ",&nbsp;" & Server.HtmlEncode(fPhraH(000802)) & ":&nbsp;" & fMin(k, i)
      If i < 99 Then
        fAssessmentStatus = fAssessmentStatus & "&nbsp;" & Server.HtmlEncode(fPhraH(000763)) & "&nbsp;" & i
      End If
      If j/100 >= fPassingScore Then 
        fAssessmentStatus = fAssessmentStatus & ", " & Server.HtmlEncode(fPhraH(000107))
      End If
    End If  
    If (svBrowser <> "msie") Then
      fAssessmentStatus = fAssessmentStatus & " | <a onclick='location.reload();' href='#'>" & Server.HtmlEncode(fPhraH(000522)) & "</a>"
    End If
  End Function  


  '...this returns the latest date that each module attained the vMinScore
  '   if any assessment was not passed then fLastPassed returns ""
  '   it is used in the SPC Certificate programs and MyContent.asp
  Function fLastPassed(vModIds, vMinScore)
    Dim aMods, vModId, i, j, k
    aMods = Split(Trim(vModIds))
    fLastPassed = "Jan 1, 2000"
    For i = 0 To Ubound(aMods)
      vModId = aMods(i)
      If fBestScore (svMembNo, vModId) < vMinScore * 100 Then 
        fLastPassed= ""
        Exit Function
      Else 
        j = fLastScore (svMembNo, vModId)
        If IsDate(j) Then
          fLastPassed = fMax(cDate(j), cDate(fLastPassed))
        Else
          fLastPassed = cDate(fLastPassed)
        End If
      End If
    Next  
  End Function      


  '...determine what type of cert we need to display, custom or basic
  '   if we are passing in a Prog Id then check that out (from status panel for ex)
  '   if it's a module then get cert info from the db (assessment report)
  '   if nothing then just check at customer level
  Function fCertFolder (vProgId)
    fCertFolder = ""
    If Len(vProgId) = 7 Then
      fCertFolder = fProgAssessmentCert (vProgId)
    ElseIf Len(vProgId) = 6 Then
      fCertFolder = fModsAssessmentCert (vProgId)
    End If
    If Len(fCertFolder) = 0 Then
      fCertFolder = fOkValue(fCustAssessmentCert (svCustId))
    End If
    If Len(fCertFolder) = 0 Then
      fCertFolder = svLang
    End If
  End Function


  '...Mod Assessment Cert?
  '   This is needed when we look from the assessment report which does not include the program id (bad original design)
  '   check all progs in the catl which use this mod id and see if there a custom cert
  Function fModsAssessmentCert (vModsId)
    fModsAssessmentCert = ""
    vSql = "SELECT TOP 1 V5_Base.dbo.Prog.Prog_AssessmentCert "_
         & "FROM V5_Vubz.dbo.Catl INNER JOIN V5_Base.dbo.Prog ON CHARINDEX(V5_Base.dbo.Prog.Prog_Id, V5_Vubz.dbo.Catl.Catl_Programs) > 0 "_
         & "WHERE (V5_Vubz.dbo.Catl.Catl_CustId = '" & svCustId & "') AND (CHARINDEX('" & vModsId & "', V5_Base.dbo.Prog.Prog_Mods) > 0) AND (LEN(V5_Base.dbo.Prog.Prog_AssessmentCert) > 0) "
    sOpenDbBase2    
    Set oRsBase2 = oDbBase2.Execute(vSql)
    If Not oRsBase2.Eof Then 
      fModsAssessmentCert = fOkValue(oRsBase2("Prog_AssessmentCert"))
    End If
    Set oRsBase2 = Nothing
    sCloseDbBase2    
  End Function


  '...Prog Assessment Cert?
  Function fProgAssessmentCert (vProgId)
    fProgAssessmentCert = ""
    vSql = "SELECT Prog_AssessmentCert FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase2    
    Set oRsBase2 = oDbBase2.Execute(vSql)
    If Not oRsBase2.Eof Then 
      fProgAssessmentCert = fOkValue(oRsBase2("Prog_AssessmentCert"))
    End If
    Set oRsBase2 = Nothing
    sCloseDbBase2    
  End Function


  '...Cust Assessment Cert?
  Function fCustAssessmentCert (vCustId)
    fCustAssessmentCert = ""
    vSql = "SELECT Cust_AssessmentCert FROM Cust WHERE Cust_Id= '" & vCustId & "'"
    sOpenDb4
    Set oRs4 = oDb4.Execute(vSql)
    If Not oRs4.Eof Then 
      fCustAssessmentCert = oRs4("Cust_AssessmentCert")
    End If
    Set oRs4 = Nothing
    sCloseDb4    
  End Function

  
  '...Get Passing Score (assumes that the prog and cust records have been read)
  '   .01 is used on the Cust/Prog tables to specify that any score is a passing/completed score
  Function fPassingScore () 
    fPassingScore = .80
    If vCust_AssessmentScore > 0 Then fPassingScore = vCust_AssessmentScore
    If vProg_AssessmentScore > 0 Then fPassingScore = vProg_AssessmentScore
  End Function
 
%>



