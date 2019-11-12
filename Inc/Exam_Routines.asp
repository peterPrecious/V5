<%
  '...functions for the Mods and Tests
   
  Function GetStrBank (vModID, vBank, vRandom)
    Dim oRs, aTemp, aQue
    Dim vQ1, vQ2, vQ3, vQ4, vQ5
    Redim aQue(4)
    '...if Random, get next 5 random questions
    If vRandom Then
    '...otherwise grab next 5 in order
      aQue = GetRandomQ(vModID, aQue)
    Else
      aQue(0) = ((vBank-1) * 5) + 1
      aQue(1) = ((vBank-1) * 5) + 2
      aQue(2) = ((vBank-1) * 5) + 3
      aQue(3) = ((vBank-1) * 5) + 4
      aQue(4) = ((vBank-1) * 5) + 5
    End If
    sOpenDbBase
    Redim aTemp(4)
    For i = 0 to 4
      vSql = "Select * FROM TstQ WHERE TstQ_ID = '" & vModID & "' AND TstQ_No = " & aQue(i)

      Set oRs = oDbBase.Execute(vSql)    
      If Not oRs.EOF Then 
'       aTemp(i) = aQue(i) & "||" & Server.HtmlEncode(oRs("TstQ_Q"))
        aTemp(i) = aQue(i) & "||" & oRs("TstQ_Q")
      End If
    Next
    GetStrBank = aTemp
    sCloseDbBase
  End Function

  Function GradeTestBank (vModID, vBank, vTestStart, vTimeLimit, aCurrentBank, vCurrentAttempt)
    Dim aQue, vQue, aRes, vStr, aAns, vAns, vFld, vValue, vTimeLen
    Dim aTemp, vQ1, vQ2, vQ3, vQ4, vQ5
    Dim vQ1desc, vQ2desc, vQ3desc, vQ4desc, vQ5desc, strQHist, oRsHist

    '...if vTimeLimit = 0, then NO time limit
    If vTimeLimit = 0 Then vTimeLimit = 99999999

    '...determine the time taken for the bank
    vTimeLen = (minute(Time - vTestStart)*60) + second(Time - vTestStart)
    '...if time was too long, give grade of zero
    If vTimeLen > (vTimeLimit*60) Then
      GradeTestBank = -999
      Exit Function
    End If
    aQue = aCurrentBank : vQue = Ubound(aQue)   
    ReDim aRes(1,vQue)
    '...get correct values
    For I = 0 To vQue
      aRes(1, i) = 0 'initialize test values
      aAns = Split(aQue(i),"||"): vAns = Ubound(aAns) 
      aRes(0, i) = aAns(2)
    Next
    '...get question numbers
    aTemp = Split(aQue(0),"||")
    vQ1 = aTemp(0)
    vQ1desc = aTemp(1)
    aTemp = Split(aQue(1),"||")
    vQ2 = aTemp(0)
    vQ2desc = aTemp(1)
    aTemp = Split(aQue(2),"||")
    vQ3 = aTemp(0)
    vQ3desc = aTemp(1)
    aTemp = Split(aQue(3),"||")
    vQ4 = aTemp(0)
    vQ4desc = aTemp(1)
    aTemp = Split(aQue(4),"||")
    vQ5 = aTemp(0)
    vQ5desc = aTemp(1)
    '...get test values
    For Each vFld in Request.Form
      vValue = Request.Form(vFld)
      Select Case vFld
        Case "Q01" : aRes(1, 0) = vValue
        Case "Q02" : aRes(1, 1) = vValue
        Case "Q03" : aRes(1, 2) = vValue
        Case "Q04" : aRes(1, 3) = vValue
        Case "Q05" : aRes(1, 4) = vValue
      End Select
    Next
    '...crib notes
    If svMembLevel > 33 Then
      Response.Write "<P><font face='Arial' size='2'>Polly: answers...<br>"
      For I = 0 To vQue - 1: Response.Write right("0" & i+1, 1): Next
      Response.write "<br>"
      For I = 0 To vQue - 1: Response.Write aRes(0, i+1): Next
      Response.write "<br>"
      For I = 0 To vQue - 1: Response.Write aRes(1, i+1): Next
      Response.Write "</P></Font>"
    End If   
    '...get mark
    j = 0 : k = 0
    For i = 0 To vQue
      '...process valid questions     
      If Not IsNumeric(aRes(0, i)) Then 
        Exit For
      Else
        k = i
        If aRes(0, i) = aRes(1,i) Then
          j = j + 1
        Else
          '...if incorrect, store the Question
          If i = 0 Then
            strQHist = strQHist & "@@" & vQ1desc
          Elseif i = 1 Then
            strQHist = strQHist & "@@" & vQ2desc
          Elseif i = 2 Then
            strQHist = strQHist & "@@" & vQ3desc
          Elseif i = 3 Then
            strQHist = strQHist & "@@" & vQ4desc
          Elseif i = 4 Then
            strQHist = strQHist & "@@" & vQ5desc
          End If
        End If
      End If
    Next
'   GradeTest = j / vQue
    If k > 0 Then
      GradeTestBank = j / (k+1)
    Else
      GradeTestBank = 0
    End If
    '...need to save results in Logs table
    Dim oRs
    sOpenDb
    '...delete zero marks for this bank
    vSql = "DELETE FROM Logs WHERE Logs_AcctID = " & svCustAcctId
    vSql = vSql & " AND Logs_MembNo = " & svMembNo
    vSql = vSql & " AND Logs_Type = 'E'"
    vSql = vSql & " AND Logs_Item LIKE '" & vModID & "_" & vCurrentAttempt & "_" & vBank & "_%'"
    oDb.Execute vSQL
    '...insert new values into Logs
    vSql = "INSERT INTO Logs (Logs_AcctID, Logs_MembNo, Logs_Type, Logs_Item, Logs_Posted) "
    vSql = vSql & "VALUES ("
    vSql = vSql & "'" & svCustAcctId & "', "
    vSql = vSql & svMembNo & ", "
    vSql = vSql & "'E', "
    vSql = vSql & "'" & vModID & "_" & vCurrentAttempt & "_" & vBank & "_" & vQ1 & "_" & aRes(1,0) & "_" & vQ2 & "_" & aRes(1,1) & "_" & vQ3 & "_" & aRes(1,2) & "_" & vQ4 & "_" & aRes(1,3) & "_" & vQ5 & "_" & aRes(1,4) & "_" & vTimeLen & "', "
    vSql = vSql & "'" & fFormatSqlDate(Now) & "'"
    vSql = vSql & ")"
    oDb.Execute vSQL
    '...store incorrect quesions answered in the Logs table under "H"
    If Len(strQHist) > 0 Then
      '...if a record already exists, we concatenate...otherwise, do initial Insert
      vSql = "SELECT * FROM Logs WHERE (Logs_AcctID='" & svCustAcctId & "' AND "
      vSql = vSql & "Logs_MembNo=" & svMembNo & " AND "
      vSql = vSql & "Logs_Type='H' AND "
      vSql = vSql & "Logs_Item LIKE '" & vModID & "_" & vCurrentAttempt & "_%')"
      Set oRsHist = oDb.Execute(vSql)
      If oRsHist.Eof Then
        '...insert all INCORRECT answered questions into Logs under "H"
        vSql = "INSERT INTO Logs (Logs_AcctID, Logs_MembNo, Logs_Type, Logs_Item, Logs_Posted) "
        vSql = vSql & "VALUES ("
        vSql = vSql & "'" & svCustAcctId & "', "
        vSql = vSql & svMembNo & ", "
        vSql = vSql & "'H', "
        vSql = vSql & "'" & vModID & "_" & vCurrentAttempt & "_" & Left(Replace(strQHist,"'","''"), 3000) & "', "
        vSql = vSql & "'" & fFormatSqlDate(Now) & "'"
        vSql = vSql & ")"
      Else
        '...update all INCORRECT answered questions into current Logs under "H"
        vSql = "UPDATE Logs SET Logs_Item = '" & Left(Replace(oRsHist("Logs_Item") & strQHist,"'","''"), 3000) & "'"
        vSql = vSql & " WHERE (Logs_AcctID='" & svCustAcctId & "' AND "
        vSql = vSql & "Logs_MembNo=" & svMembNo & " AND "
        vSql = vSql & "Logs_Type='H' AND "
        vSql = vSql & "Logs_Item LIKE '" & vModID & "_" & vCurrentAttempt & "_%')"
      End If      
      oRsHist.Close
      Set oRsHist = Nothing
      oDb.Execute vSQL
    End If
    sCloseDb
  End Function

  Function GradeInitBank (vModID, vBank, aQue, vCurrentAttempt)
    '...need to save results in Logs table
    Dim oRs, aTemp, vQ1, vQ2, vQ3, vQ4, vQ5
    GradeInitBank = True
    sOpenDb
    '...first check to see if entry already there
    vSQL = "SELECT * FROM Logs WHERE (Logs_AcctID = " & svCustAcctId
    vSQL = vSQL & " AND Logs_MembNo = " & svMembNo
    vSQL = vSQL & " AND Logs_Type = 'E'"
    vSQL = vSQL & " AND Logs_Item LIKE '" & vModID & "_" & vCurrentAttempt & "_" & vBank & "_%')"
    Set oRS = oDb.Execute(vSQL)
    If Not oRS.EOF Then
      GradeInitBank = False
      Exit Function
    End If
    '...get question numbers
    aTemp = Split(aQue(0),"||")
    vQ1 = aTemp(0)
    aTemp = Split(aQue(1),"||")
    vQ2 = aTemp(0)
    aTemp = Split(aQue(2),"||")
    vQ3 = aTemp(0)
    aTemp = Split(aQue(3),"||")
    vQ4 = aTemp(0)
    aTemp = Split(aQue(4),"||")
    vQ5 = aTemp(0)
    vSql = "INSERT INTO Logs (Logs_AcctID, Logs_MembNo, Logs_Type, Logs_Item, Logs_Posted) "
    vSql = vSql & "VALUES ("
    vSql = vSql & "'" & svCustAcctId & "', "
    vSql = vSql & svMembNo & ", "
    vSql = vSql & "'E', "
    vSql = vSql & "'" & vModID & "_" & vCurrentAttempt & "_" & vBank & "_" & vQ1 & "_0_" & vQ2 & "_0_" & vQ3 & "_0_" & vQ4 & "_0_" & vQ5 & "_0_0', "
    vSql = vSql & "'" & fFormatSqlDate(Now) & "'"
    vSql = vSql & ")"
    oDb.Execute vSQL
    sCloseDb
  End Function

  Function GetTotalResults (vModID, vTotalTime, vBankCurrent)
    Dim aQue, vQue, aRes, vStr, aAns, vAns, vFld, vValue, vTimeLen
    Dim vBank, aResults, vNumBanks, vGradeBank, vAttempt
    GetTotalResults = 0
    vTotalTime = 0
    '...get info on all saved banks to date
    TestInProgress vModID, vBank, aResults, vAttempt
    For vNumBanks = 0 To (vBankCurrent-1)
      aQue = Split(aResults(vNumBanks),"_") : vQue = UBound(aQue)   
      ReDim aRes(1, 4)
      '...get correct values
      aRes(0, 0) = GetCorrectValue(aQue(3))
      aRes(0, 1) = GetCorrectValue(aQue(5))
      aRes(0, 2) = GetCorrectValue(aQue(7))
      aRes(0, 3) = GetCorrectValue(aQue(9))
      aRes(0, 4) = GetCorrectValue(aQue(11))
      '...get test values
      aRes(1, 0) = aQue(4)
      aRes(1, 1) = aQue(6)
      aRes(1, 2) = aQue(8)
      aRes(1, 3) = aQue(10)
      aRes(1, 4) = aQue(12)
      '...get mark
      j = 0 : k = 0
      For i = 0 To 4
        '...process valid questions     
        If Not IsNumeric(aRes(0, i)) Then 
          Exit For
        Else
          k = i
          If aRes(0, i) = aRes(1,i) Then j = j + 1    
        End If
      Next
      If k > 0 Then
        'vGradeBank = j / (k+1)
        vGradeBank = j
      Else
        vGradeBank = 0
      End If
      GetTotalResults = GetTotalResults + vGradeBank
      vTotalTime = vTotalTime + aQue(13)
    Next
    GetTotalResults = GetTotalResults / (vBankCurrent * 5)
  End Function

  Function TestInProgress (vModID, vBank, aResults, vCurrentAttempt)
    Dim i, aTemp, aTempQ, vExpires
    TestInProgress = False
    vCurrentAttempt = GetNumberAttempts(vModID)
    If vCurrentAttempt = 0 Then Exit Function
    If vBank = 1 Then Exit Function '...if on first bank, no need to continue   

    vExpires = DateAdd("yyyy", -1, Now) '...used to isolate only current exam info

    sOpenDb

    vSQL = "SELECT * FROM Logs WHERE (Logs_AcctID = " & svCustAcctId
    vSQL = vSQL & " AND Logs_MembNo = " & svMembNo
'   vSQL = vSQL & " AND Logs_Type = 'E'"
    vSQL = vSQL & " AND Logs_Type = 'E' AND Logs_Posted > '"  & vExpires & "'"
    vSQL = vSQL & " AND Logs_Item LIKE '" & vModID & "_" & vCurrentAttempt & "_%')"

    Set oRS = oDb.Execute(vSQL)
    If oRS.EOF Then Exit Function
    TestInProgress = True
    '...get the test results, etc. for each bank
    ReDim aTemp (0)
    i = -1
    While Not oRS.EOF
      i = i + 1
      aTemp(i) = oRS("Logs_Item")
      oRS.MoveNext
      ReDim Preserve aTemp(UBound(aTemp)+1)
    Wend
    ReDim Preserve aTemp(UBound(aTemp)-1)
    '...get the last bank completed
'    vBank = Mid(aTemp(UBound(aTemp)),8,1)
    aTempQ = Split(aTemp(UBound(aTemp)),"_")
    vBank = aTempQ(2)
    aResults = aTemp   
    Set oRS = Nothing
    sCloseDb
  End Function

  Function GetCorrectValue (vQ)
    Dim aQue
    sOpenDbBase
    vSql = "Select * FROM TstQ WHERE TstQ_ID = '" & vModID & "' AND TstQ_No = " & vQ
    Set oRs = oDbBase.Execute(vSQL)    
    aQue = Split(oRs("TstQ_Q"),"||") 
    GetCorrectValue = aQue(1)
    sCloseDbBase
  End Function

  Function GetRandomQ(vModID, aQue)
    Dim vQCount, vQCheck, vOK, vNum
    Dim aResults, vMaxQue, vBank, vAttempt
    Dim aExistQue, aTemp, vFindCount
    '...define local vars
    vMaxQue = Session(vModID & "MaxQue")
    vBank = Session(vModID & "Bank")
    vAttempt = Session(vModID & "Attempt")
    '...get info on all saved banks to date
    TestInProgress vModID, vBank, aResults, vAttempt
    '...split out all questions into a 2 dim array (bank,que)
    If VarType(aResults) = vbEmpty Then
      ReDim aExistQue (-1, 4)
    Else
      ReDim aExistQue (UBound(aResults), 4)
      For i = 0 To Ubound(aExistQue)
        aTemp = Split(aResults(i),"_")
        aExistQue(i,0) = CInt(aTemp(3))
        aExistQue(i,1) = CInt(aTemp(5))
        aExistQue(i,2) = CInt(aTemp(7))
        aExistQue(i,3) = CInt(aTemp(9))
        aExistQue(i,4) = CInt(aTemp(11))
      Next
    End If
    Randomize Timer
    For vQCount = 0 to 4
      vOK = False
      While Not vOK
        vOK = True
        '...use MaxQue to determine what to choose from
        vNum = Int(vMaxQue * Rnd + 1)
        '...make sure number not already generated within THIS bank
        For vQCheck = 0 To vQCount-1
          If aQue(vQCheck) = vNum Then
            vOK = False
            Exit For
          End If
        Next
        '...make sure number not already within ANY of the banks
        If vOK And UBound(aExistQue,1) >= 0 Then
          For vQCheck = 0 To UBound(aExistQue,1)
            If aExistQue(vQCheck,0) = vNum OR aExistQue(vQCheck,1) = vNum OR aExistQue(vQCheck,2) = vNum OR aExistQue(vQCheck,3) = vNum OR aExistQue(vQCheck,4) = vNum Then
              vOK = False
              Exit For
            End If
          Next
        End If
      WEnd
      aQue(vQCount) = vNum
    Next
    GetRandomQ = aQue
  End Function

  Function GetRandomQ_BACKUP_CONDITIONAL_RANDOM(vModID, aQue)
    Dim vQCount, vQCheck, vOK, vNum
    Dim aResults, vMaxQue, vBank, vAttempt
    Dim aExistQue, aTemp, vFindCount
    '...define local vars
    vMaxQue = Session(vModID & "MaxQue")
    vBank = Session(vModID & "Bank")
    vAttempt = Session(vModID & "Attempt")
    '...get info on all saved banks to date
    TestInProgress vModID, vBank, aResults, vAttempt
    '...split out all questions into a 2 dim array (bank,que)
    If VarType(aResults) = vbEmpty Then
      ReDim aExistQue (-1, 4)
    Else
      ReDim aExistQue (UBound(aResults), 4)
      For i = 0 To Ubound(aExistQue)
        aTemp = Split(aResults(i),"_")
        aExistQue(i,0) = CInt(aTemp(3))
        aExistQue(i,1) = CInt(aTemp(5))
        aExistQue(i,2) = CInt(aTemp(7))
        aExistQue(i,3) = CInt(aTemp(9))
        aExistQue(i,4) = CInt(aTemp(11))
      Next
    End If
    Randomize Timer
    For vQCount = 0 to 4
      vOK = False
      While Not vOK
        vOK = True
        '...use MaxQue to determine what to choose from
        vNum = Int(vMaxQue * Rnd + 1)
        '...make sure number not already generated within THIS bank
        For vQCheck = 0 To vQCount-1
          If aQue(vQCheck) = vNum Then
            vOK = False
            Exit For
          End If
        Next
        '...make sure number not already within last ONE banks
        If vOK And UBound(aExistQue,1) >= 0 Then
          For vQCheck = UBound(aExistQue,1) To UBound(aExistQue,1)
            If aExistQue(vQCheck,0) = vNum OR aExistQue(vQCheck,1) = vNum OR aExistQue(vQCheck,2) = vNum OR aExistQue(vQCheck,3) = vNum OR aExistQue(vQCheck,4) = vNum Then
              vOK = False
              Exit For
            End If
          Next
        End If
        '...make sure number not already within last TWO banks
        If vOK And UBound(aExistQue,1) >= 1 Then
          For vQCheck = UBound(aExistQue,1)-1 To UBound(aExistQue,1)-1
            If aExistQue(vQCheck,0) = vNum OR aExistQue(vQCheck,1) = vNum OR aExistQue(vQCheck,2) = vNum OR aExistQue(vQCheck,3) = vNum OR aExistQue(vQCheck,4) = vNum Then
              vOK = False
              Exit For
            End If
          Next
        End If
        '...make sure number generated more than once within THIS exam
        If vOK And UBound(aExistQue,1) >= 2 Then
          vFindCount = 0
          For vQCheck = 0 To UBound(aExistQue,1)-2
            If aExistQue(vQCheck,0) = vNum OR aExistQue(vQCheck,1) = vNum OR aExistQue(vQCheck,2) = vNum OR aExistQue(vQCheck,3) = vNum OR aExistQue(vQCheck,4) = vNum Then
              vFindCount = vFindCount + 1
              If vFindCount = 2 Then
                vOK = False
                Exit For
              End If
            End If
          Next
        End If
      WEnd
      aQue(vQCount) = vNum
    Next
    GetRandomQ = aQue
  End Function

  Function DeleteScores(vModID)
    Dim oRs
    sOpenDb
    '...delete marks...for debugging ONLY
    vSql = "DELETE FROM Logs WHERE Logs_AcctID = " & svCustAcctId
    vSql = vSql & " AND Logs_MembNo = " & svMembNo
    vSql = vSql & " AND Logs_Type = 'E'"
    vSql = vSql & " AND Logs_Item LIKE '" & vModID & "_%'"
    oDb.Execute vSQL
    sCloseDb
  End Function 

  Function GetExamTitle(vModID)
    sOpenDbBase
    vSql = "Select * FROM TstH WHERE TstH_ID = '" & vModID & "'"
    Set oRs = oDbBase.Execute(vSql)    
    If Not oRs.Eof Then GetExamTitle = oRs("TstH_Title")
    sCloseDbBase
  End Function

  Function GetNumberAttempts(vModID)
    Dim aTemp
    sOpenDb
    vSQL = "SELECT * FROM Logs WHERE (Logs_AcctID = " & svCustAcctId    
    vSQL = vSQL & " AND Logs_MembNo = " & svMembNo
    vSQL = vSQL & " AND Logs_Type = 'E'"
    vSQL = vSQL & " AND Logs_Item LIKE '" & vModID & "_%')"
    vSQL = vSQL & " ORDER BY Logs_Item DESC"
    Set oRS = oDb.Execute(vSQL)
    If oRS.EOF Then
      GetNumberAttempts = 0
    Else
      aTemp = Split(oRs("Logs_Item"),"_")
      GetNumberAttempts = aTemp(1)
    End If
    sCloseDb
  End Function

  Function GetMaxQue(vModID)
    Dim aTemp
    sOpenDbBase
    vSQL = "SELECT * FROM TstH WHERE TstH_ID = '" & vModID & "'"
'   sDebug
    Set oRS = oDbBase.Execute(vSQL)
    GetMaxQue = oRS("TstH_NoQuestions")
    sCloseDbBase
  End Function

  Function GetQuestion(vModID, vQNum)
    Dim aTemp, oRS
    vSQL = "SELECT * FROM TstQ WHERE TstQ_ID = '" & vModID & "' AND TstQ_No = " & vQNum
    Set oRS = oDbBase.Execute(vSQL)
    aTemp = Split(oRS("TstQ_Q"),"||")
    GetQuestion = aTemp(0)
  End Function

%>