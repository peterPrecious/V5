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

  Function GetStrBankAll (vModID)
    Dim oRs, aTemp, aQue
    Dim vQ1, vQ2, vQ3, vQ4, vQ5
    sOpenDbBase
    vSql = "Select * FROM TstQ WHERE TstQ_ID = '" & vModID & "'"
    Set oRs = oDbBase.Execute(vSql)    
    Redim aTemp(-1)
    While Not oRs.EOF
      Redim Preserve aTemp(UBound(aTemp)+1)
'     aTemp(UBound(aTemp)) = Server.HtmlEncode(oRs("TstQ_Q"))
      aTemp(UBound(aTemp)) = oRs("TstQ_Q")
      oRs.MoveNext
    Wend
    sCloseDbBase
    GetStrBankAll = aTemp
  End Function

  Function GetStrBankEdit (vModID, vBank)
    Dim oRs, aTemp, aQue
    Dim vQ1, vQ2, vQ3, vQ4, vQ5
    Redim aQue(4)
    aQue(0) = ((vBank-1) * 5) + 1
    aQue(1) = ((vBank-1) * 5) + 2
    aQue(2) = ((vBank-1) * 5) + 3
    aQue(3) = ((vBank-1) * 5) + 4
    aQue(4) = ((vBank-1) * 5) + 5
    sOpenDbBase
    Redim aTemp(4)
    For i = 0 to 4
      vSql = "Select * FROM TstQ WHERE TstQ_ID = '" & vModID & "' AND TstQ_No = " & aQue(i)
      Set oRs = oDbBase.Execute(vSql)    
      If Not oRs.EOF Then 
        aTemp(i) = aQue(i) & "||" & Server.HtmlEncode(oRs("TstQ_Q"))
      Else
        aTemp(i) = aQue(i) & "||" & "||1||||||||||||||"
      End If
    Next
    GetStrBankEdit = aTemp
    sCloseDbBase
  End Function

  Sub SaveQuestionsBank(vBank)
    Dim vFld, vValue, vStr, vSql, vQue
    '...get Mod ID
    vModID = Request.Form("vModID")    
    '...get Question info
    vQue = Session("Que")
      
    sOpenDbBase

    '...if Bank 1, save Exam Title
    If vBank = 1 Then
      vSql = "UPDATE TstH SET"
      vSql = vSql & " TstH_Title = '" & fUnquote(Request.Form("vTitle")) & "'"
      vSql = vSql & " WHERE TstH_ID = '" & vModID & "'"

      oDbBase.Execute vSQL
    End If

    '...build string      
    For i =  0 to 4
      '...delete question for this bank
      vSql = "DELETE FROM TstQ WHERE TstQ_ID='" & vModID & "' AND TstQ_No=" & ((vBank-1)*5) + i+1
      oDbBase.Execute vSQL

      '...save Test
      vSql = "INSERT INTO TstQ (TstQ_ID, TstQ_No, TstQ_Q) "
      vSql = vSql & "VALUES ("
      vSql = vSql & "'" & vModID & "', "
      vSql = vSql & ((vBank-1)*5) + i+1 & ", "
      vSql = vSql & "'" & fUnquote(Mid(vQue(i),InStr(vQue(i),"||")+2)) & "'"
      vSql = vSql & ")"
      oDbBase.Execute vSQL
    Next

    sCloseDbBase

  End Sub  

  Function AddExam()
    Dim vMod, vNumQ, vTitle

    On Error Resume Next

    '...get Info
    vMod = UCase(Request.Form("vAddModID"))
    vNumQ = Request.Form("vNumQ")
    vTitle = Request.Form("vTitle")
    
    AddExam = ""

    '...validate info
    If Len(vMod) <> 6 Then
      AddExam = "Exam name must be 6 characters."
      Exit Function
    ElseIf Not IsNumeric(Left(vMod,4)) Then
      AddExam = "First 4 characters of the Exam Name must be numeric."
      Exit Function
    ElseIf (Right(vMod,2) <> "EN") And (Right(vMod,2) <> "FR") And (Right(vMod,2) <> "ES") And (Right(vMod,2) <> "PT") Then
      AddExam = "Last 2 characters of the Exam Name must be EN, FR, ES or PT."
      Exit Function
    ElseIf Not IsNumeric(vNumQ) Then
      AddExam = "Number of Questions must be numeric."
      Exit Function
    ElseIf CInt(vNumQ) < 5 Then
      AddExam = "Number of Questions must be at least 5."
      Exit Function
    ElseIf (CInt(vNumQ) Mod 5) <> 0 Then
      AddExam = "Number of Questions must be in increments of 5."
      Exit Function
    ElseIf Len(vTitle) = 0 Then
      AddExam = "Exam must have a Title."
      Exit Function
    ElseIf Len(vTitle) > 64 Then
      AddExam = "Exam Title can be no greater than 64 characters."
      Exit Function
    End If

    '...build string      
    sOpenDbBase

    vSql = "INSERT INTO TstH (TstH_ID, TstH_NoQuestions, TstH_Title) "
    vSql = vSql & "VALUES ("
    vSql = vSql & "'" & vMod & "', "
    vSql = vSql & vNumQ & ", "
    vSql = vSql & "'" & fUnquote(vTitle) & "'"
    vSql = vSql & ")"
    oDbBase.Execute vSQL

    sCloseDbBase

    If Err = -2147217900 Then
      AddExam = "Exam " & vMod & " already exists; cannot duplicate."
    ElseIf Err <> 0 Then
      AddExam = Err & ":  " & Err.Description
    Else
      AddExam = "Exam " & vMod & " has been added successfully."
    End If

  End Function

  Function GetExamTitle(vModID)
    sOpenDbBase
    vSql = "Select * FROM TstH WHERE TstH_ID = '" & vModID & "'"
    Set oRs = oDbBase.Execute(vSql)    
    GetExamTitle = oRs("TstH_Title")
    sCloseDbBase
  End Function

  Function ValidateQuestionsBank(vMess, vBank)
    Dim vQue(7,5), vAnsCount
    ValidateQuestionsBank = True
    '...if Bank 1, check if valid Exam Title
    If vBank = 1 Then
      If Len(Request.Form("vTitle")) = 0 Then
        vMess = "Exam Title cannot be blank."
        ValidateQuestionsBank = False
        Exit Function
      ElseIf Len(Request.Form("vTitle")) > 64 Then
        vMess = "Exam Title cannot longer than 64 characters."
        ValidateQuestionsBank = False
        Exit Function
      End If
    End If

    '...get test values from edit form
    For Each vFld in Request.Form
      vValue = Request.Form(vFld)

      '...store question Qnnn in vQue(0,nnn)
      If Left(vFld,1) = "Q" and Len(vFld) = 4 Then
        i = Cint(Right(vFld,3)) - ((vBank-1)*5)
        vQue(0, i) = vValue
      End If

      '...store possible answers Annnx in vQue(1+x,nnn)
      If Left(vFld,1) = "Q" and Len(vFld) = 5 Then
        i = Cint(Mid(vFld,2,3)) - ((vBank-1)*5)
        j = instr("ABCDEF", Right(vFld,1)) + 1
        vQue(j, i) = vValue
      End If

      '...store answers Annn in vQue(1,nnn)
      If Left(vFld,1) = "A" Then
        i = Cint(Right(vFld,3)) - ((vBank-1)*5)
        vQue(1, i) = vValue
      End If
    Next            

    '...Check that all answers there
    For i = 1 to 5
      If Len(vQue(0,i)) = 0 Then
        vMess = "Please ensure that all Questions are defined."
        ValidateQuestionsBank = False
        Exit For
      End If
    Next
    '...Check that there at least 2 answers/question
    If ValidateQuestionsBank Then
      For i = 1 to 5
        vAnsCount = 0
        For j = 2 to 7
          If Len(vQue(j,i)) > 0 Then
            vAnsCount = vAnsCount + 1
          End If
        Next
        If vAnsCount < 2 Then
          vMess = "Please ensure that all Questions have at least 2 Answers defined."
          ValidateQuestionsBank = False
          Exit For
        End If
      Next
    End If

    '...Store the Bank questions
    Dim aTemp(4)
    For i =  0 to 4
      aTemp(i) = ((vBank-1)*5) + (i+1)
      For j = 0 to 7
        aTemp(i) = aTemp(i) & "||" & vQue(j,i+1)
      Next
    Next
    Session("Que") = aTemp

  End Function

  Function sDeleteExam (vModId)

    On Error Resume Next
    sDeleteExam = True

    sOpenDbBase
    oDbBase.BeginTrans

    vSql = "DELETE TstQ WHERE TstQ_ID = '" & vModId & "'"
    oDbBase.Execute(vSql)

    If Err.Number <> 0 Then
      sDeleteExam = False
      oDbBase.RollbackTrans
    Else
      vSql = "DELETE TstH WHERE TstH_ID = '" & vModId & "'"
      oDbBase.Execute(vSql)
      If Err.Number <> 0 Then
        sDeleteExam = False
        oDbBase.RollbackTrans
      Else
        oDbBase.CommitTrans
      End If
    End If

    sCloseDbBase
  End Function

%>