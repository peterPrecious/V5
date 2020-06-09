<%
  '...functions for the Mods and Tests

  Sub sDeleteTest (vModId)
    vSql = "DELETE Test WHERE Test_Mod = '" & vModId & "'"
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub
   
  Function fModsOK (vModId)
    fModsOK = False
    If Instr(vModIds, vModId) > 0 Then fModIdOK = True
  End Function

  Function fModsTitle (vModId)
    fModsTitle = ""
    sOpenDbBase
    vSQL = "SELECT * FROM Mods WHERE Mods_Id = '" & vModId & "'" 
    Set oRS = oDbBase.Execute(vSQL)
    If oRS.Eof Then Exit Function
    fModsTitle = oRS("Mods_Title")
    Set oRS = Nothing
    sCloseDbBase
  End Function

  Function fExamTitle (vTstHId)
    fExamTitle = ""
    sOpenDbBase
    vSQL = "SELECT * FROM TstH WHERE TstH_Id = '" & vTstHId& "'" 
    Set oRS = oDbBase.Execute(vSQL)
    If oRS.Eof Then Exit Function
    fExamTitle = oRS("TstH_Title")
    Set oRS = Nothing
    sCloseDbBase
  End Function

  Function fModsDesc (vModId)
    fModsDesc = ""
    sOpenDbBase
    vSQL = "SELECT * FROM Mods WHERE Mods_Id = '" & vModId & "'" 
    Set oRS = oDbBase.Execute(vSQL)
    If oRS.Eof Then Exit Function
    fModsDesc = oRS("Mods_Desc")
    Set oRS = Nothing
    sCloseDbBase
  End Function

  Function fModsOneHour (vModId)
    fModsOneHour = False
    sOpenDbBase
    vSQL = "SELECT Mods_Cat FROM Mods WHERE Mods_Id = '" & vModId & "'" 
    Set oRS = oDbBase.Execute(vSQL)
    If oRS.Eof Then Exit Function
    If Instr("00 01", oRS("Mods_Cat")) > 0 then fModsOneHour = True
    Set oRS = Nothing
    sCloseDbBase
  End Function
  
  Sub InitTest (vModId)
    Dim vSql
    sOpenDbBase  
    vSql = "INSERT INTO Test " _
         & "(Test_Mod, Test_Str) VALUES " _
         & "('" & vModId & "', '" & InitStr & "')"
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub   
  
  Function GetStr (vModId)
    Dim oRs
    sOpenDbBase
    vSql = "Select * FROM Test WHERE Test_Mod = '" & vModId & "'" 
'   response.write "<P>" & vSQl    
    Set oRs = oDbBase.Execute(vSQL)    
    If oRs.Eof Then 
      sCloseDbBase
      InitTest vModId
      GetStr = InitStr
    Else
      GetStr = Server.HtmlEncode(oRs("Test_Str"))
      sCloseDbBase
    End If        
'   Response.write "<P>Str: " & GeTestr
  End Function
  
  '...determine if Test Active
  Function fTestActive (vModId)
    Dim oRs, vCou_Test
    fTestActive = False
    sOpenDbBase
    vSql = "Select Mods_Test FROM Mods WHERE Mods_Id = '" & vModId & "'"
    Set oRs = oDbBase.Execute(vSQL)    
    If Not oRs.Eof Then 
      vCou_Test = oRs("Mods_Test")
      If Not fNoValue(vCou_Test) And vCou_Test = True Then 
        fTestActive = True
      End If
    End If
    sCloseDbBase        
  End Function
  
    '...get eligible testsCloseDB from database
  Function fTestOptionsAll
    Dim oRs
    fTestOptionsAll = ""
    sOpenDbBase
    vSql = "Select * FROM Test ORDER by Test_Mod"
    Set oRs = oDbBase.Execute(vSQL)    
    Do While NOT oRs.Eof 
      fTestOptionsAll = fTestOptionsAll & "<option>" & oRs("Test_Mod") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop      
    sCloseDbBase           
  End Function


  Function fExamOptionsAll
    Dim oRs
    fExamOptionsAll = ""
    sOpenDbBase
    vSql = "Select * FROM TstH" 
    Set oRs = oDbBase.Execute(vSQL)    
    Do While NOT oRs.Eof 
      fExamOptionsAll = fExamOptionsAll & "<option>" & oRs("TstH_Id") & "</option>"
      oRs.MoveNext
    Loop      
    sCloseDbBase           
  End Function

  
  Function fTestOptionsActive
    Dim oRs
    fTestOptionsActive = ""
    sOpenDbBase
    vSql = "Select * FROM Mods WHERE Mods_Test ORDER by Mods_Id" 
    Set oRs = oDbBase.Execute(vSQL)    
    Do While NOT oRs.Eof 
      fTestOptionsActive = fTestOptionsActive & "<option>" & oRs("Mods_Id") & " - " & oRs("Mods_Title") & "</option>"
      oRs.MoveNext
    Loop      
    sCloseDbBase           
  End Function

 
  Function InitStr
    Dim i
    InitStr = ""
    For i = 1 To 20
      InitStr = InitStr & "||||||||||||||~~"
    Next
  End Function
  
  Function GradeTest (vModId)
    Dim aQue, vQue, aRes, vStr, aAns, vAns, vFld, vValue

    vStr = GetStr (vModId)
    aQue = Split(vStr,"~~"): vQue = Ubound(aQue)
    ReDim aRes(1,vQue)

    '...get correct values
    For I = 0 To vQue - 1
      aRes(1, i+1) = 0 'initialize test values
      aAns = Split(aQue(i),"||"): vAns = Ubound(aAns) 
      aRes(0, i+1) = aAns(1)
    Next

    '...get test values
    For Each vFld in Request.Form
      vValue = Request.Form(vFld)
      Select Case vFld
        Case "Q01" : aRes(1, 01) = vValue
        Case "Q02" : aRes(1, 02) = vValue
        Case "Q03" : aRes(1, 03) = vValue
        Case "Q04" : aRes(1, 04) = vValue
        Case "Q05" : aRes(1, 05) = vValue
        Case "Q06" : aRes(1, 06) = vValue
        Case "Q07" : aRes(1, 07) = vValue
        Case "Q08" : aRes(1, 08) = vValue
        Case "Q09" : aRes(1, 09) = vValue
        Case "Q10" : aRes(1, 10) = vValue
        Case "Q11" : aRes(1, 11) = vValue
        Case "Q12" : aRes(1, 12) = vValue
        Case "Q13" : aRes(1, 13) = vValue
        Case "Q14" : aRes(1, 14) = vValue
        Case "Q15" : aRes(1, 15) = vValue
        Case "Q16" : aRes(1, 16) = vValue
        Case "Q17" : aRes(1, 17) = vValue
        Case "Q18" : aRes(1, 18) = vValue
        Case "Q19" : aRes(1, 19) = vValue
        Case "Q20" : aRes(1, 20) = vValue    
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
    For i = 1 To vQue
      '...process valid questions     
      If Not IsNumeric(aRes(0, i)) Then 
        Exit For
      Else
        k = i
        If aRes(0, i) = aRes(1,i) Then j = j + 1    
      End If
    Next
'   GradeTest = j / vQue
    If k > 0 Then
      GradeTest = j / k
    Else
      GradeTest = 0
    End If

  End Function
  
  Sub sSaveQuestions    
    Dim vFld, vValue, vQue(7, 20), vStr
       
    '...get Mod Id
    vModId = Request.Form("vModId")    
       
    '...get test values from edit form
    For Each vFld in Request.Form
      vValue = Request.Form(vFld)
'     Response.write "<br>vFld: " & vFld & " - " & vValue

      '...store question Qnn in vQue(0,nn)
      If Left(vFld,1) = "Q" and Len(vFld) = 3 Then
        i = Cint(Right(vFld,2))
        vQue(0, i) = vValue
      End If

      '...store possible answers Annx in vQue(1+x,nn)
      If Left(vFld,1) = "Q" and Len(vFld) = 4 Then
        i = Cint(Mid(vFld,2,2))
        j = instr("ABCDEF", Right(vFld,1)) + 1
        vQue(j, i) = vValue
      End If

      '...store answers Ann in vQue(1,nn)
      If Left(vFld,1) = "A" Then
        i = Cint(Right(vFld,2))
        vQue(1, i) = vValue
      End If
    next            
      
    '...build string      
    vStr = ""
    For i = 1 To 20
      For j = 0 To 7
        vStr = vStr & vQue(j, i) & "||"
'       Response.write "<br>" & i & ", " & j & " - " & vQue(j, i)
      Next     
      vStr = vStr &  "~~"
'     Response.write "<br>" & vStr
    Next   
'   Response.write "<p>" & vStr    

    '...Save Test
    SaveTest vStr

  End Sub  
  
  Sub SaveTest (vStr)
    Dim vSql  
    sOpenDbBase
    vSql = "UPDATE Test SET"
    vSql = vSql & "  Test_Str = '" & Replace(vStr,"'","''") & "' " 
    vSql = vSql & " WHERE Test_Mod = '" & vModId & "'" 
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub     

%>