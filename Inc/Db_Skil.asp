

<%
  Dim vSkil_AcctId, vSkil_No, vSkil_Id, vSkil_Desc
  Dim vSkil_Eof

 
  '...Get all Skills
  Sub sGetSkil_Rs
    vSql = "SELECT * FROM Skil WHERE Skil_AcctId = '" & svCustAcctId & "'"
'   sDebug
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vSql)
  End Sub

  '...Get Skil Record
  Sub sGetSkil
    vSkil_Eof = False
    vSql = "SELECT * FROM Skil WHERE Skil_Id= '" & vSkil_Id & "'"
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vSql)
    If Not oRs4.Eof Then 
      sReadSkil
      vSkil_Eof = True
    End If
    Set oRs4 = Nothing
    sCloseDb4
  End Sub

  Sub sReadSkil
    vSkil_No      = oRs4("Skil_No")
    vSkil_Id      = oRs4("Skil_Id")
    vSkil_Desc    = oRs4("Skil_Desc")
  End Sub

  Sub sExtractSkil
    vSkil_Id      = Request.Form("vSkil_Id")
    vSkil_Desc    = Request.Form("vSkil_Desc")
  End Sub
  
  Sub sInsertSkil
    vSql = "INSERT INTO Skil "
    vSql = vSql & "(Skil_AcctId, Skil_Id, Skil_Desc)"
    vSql = vSql & " VALUES ('" & svCustAcctId & "', '" & vSkil_Id & "', '" & fUnquote(vSkil_Desc) & "')"
'   sDebug
    On Error Resume Next
    vFileOK = False   
    sOpenDb4
    oDb4.Execute(vSql)
    If Err.Number = 0 or Err.Number = "" Then 
      vFileOk = True
    Else
      vFileDesc = Err.Description
    End If
    On Error GoTo 0
    sCloseDb4
  End Sub

  Sub sUpdateSkil
    vSql = "UPDATE Skil SET"
    vSql = vSql & " Skil_Desc    = '" & vSkil_Desc      & "' " 
    vSql = vSql & " WHERE Skil_Id = '" & vSkil_Id & "' And Skil_AcctId = '" & svCustAcctId & "'"
    sOpenDb4
'   sDebug
    oDb4.Execute(vSql)
    sCloseDb4
  End Sub
  
  Sub sDeleteSkil
    vSql = "DELETE FROM Skil WHERE Skil_Id = '" & vSkil_Id & "' And Skil_AcctId = '" & svCustAcctId & "'"
    sOpenDb4
    oDb4.Execute(vSql)
    sCloseDb4
  End Sub

  '...get all Skills, highlight vId
  Function fSkilOptions (vId)
    Dim vSelected
    fSkilOptions = "<option>Select Skill</option>"
    sGetSkil_rs   
    Do While Not oRs4.Eof 
      vSelected = ""
      If Instr(vId, oRs4("Skil_Id")) > 0 Then 
        vSelected = " selected" 
      End If
      fSkilOptions = fSkilOptions & "<option value=" & Chr(34) & oRs4("Skil_Id") & Chr(34) & vSelected & ">" & oRs4("Skil_Id") & "</option>" & vbCrLf
      oRs4.MoveNext
    Loop      
    sCloseDb4           
  End Function

  '...get all Skill Ratung, highlight vId
  Function fSkilRateOptions (vNo)
    Dim i, vSelected
    fSkilRateOptions = ""
    fSkilRateOptions = fSkilRateOptions  & "<option value='0'>Rating</option>" & vbCrLf
    For i = 1 to 8
      vSelected = ""
      If Cint(vNo) = i Then vSelected = " selected" 
      fSkilRateOptions = fSkilRateOptions  & "<option value='" & i & "'" & vSelected & ">" & i & "</option>" & vbCrLf
    Next    
  End Function


  '...return all Skills as a string
  Function fSkills ()
    fSkills = ""
    sGetSkil_rs
    Do While Not oRs4.Eof
     sReadSkil
     fSkills = fSkills & vSkil_Id & "~"
     oRs4.MoveNext
    Loop
    Set oRs4 = Nothing
    sCloseDb4
    If Right(fSkills, 1) = "~" Then fSkills = Left(fSkills, Len(fSkills)-1)    
  End Function

%>