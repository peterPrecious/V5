<%
  Dim vResc_No, vResc_AcctId, vResc_MembNo, vResc_ResTNo, vResc_ToName, vResc_ToEmail, vResc_Intro, vResc_Mods, vResc_Date, vResc_DateSent, vResc_LastDate, vResc_NoVisits, vResc_Active
  Dim vResc_FromName, vResc_FromEmail, vResc_Subject, vResc_Body, vResc_Note '...these are extracted from the Template fields

  Dim vResc_Eof

  Sub sReadResc
    vResc_AcctId        = oRs("Resc_AcctId")
    vResc_No            = oRs("Resc_No")
    vResc_ResTNo        = oRs("Resc_ResTNo")
    vResc_MembNo        = oRs("Resc_MembNo")
    vResc_ToName        = oRs("Resc_ToName")
    vResc_ToEmail       = oRs("Resc_ToEmail")
    vResc_Intro         = oRs("Resc_Intro")
    vResc_Mods          = oRs("Resc_Mods")
    vResc_Date          = oRs("Resc_Date")       '...date created, might only have previewed, datesent added when email is sent successfully
    vResc_DateSent      = oRs("Resc_DateSent")
    vResc_LastDate      = oRs("Resc_LastDate")
    vResc_NoVisits      = oRs("Resc_NoVisits")
    vResc_Active        = oRs("Resc_Active")

    vResc_FromName      = oRs("Resc_FromName")
    vResc_FromEmail     = oRs("Resc_FromEmail")
    vResc_Subject       = oRs("Resc_Subject")
    vResc_Body          = oRs("Resc_Body")
    vResc_Note          = oRs("Resc_Note")

  End Sub


  Sub sExtractResc
    vResc_No            = Request("vResc_No")
    vResc_ResTNo        = Request("vResc_ResTNo")
    vResc_ToName        = fUnquote(Request("vResc_ToName"))
    vResc_ToEmail       = fUnquote(Request("vResc_ToEmail"))
    vResc_Intro         = fUnquote(Request("vResc_Intro"))
    vResc_Mods          = Request("vResc_Mods")

    vResc_FromName      = fUnquote(Request("vResc_FromName"))
    vResc_FromEmail     = fUnquote(Request("vResc_FromEmail"))
    vResc_Subject       = fUnquote(Request("vResc_Subject"))
    vResc_Body          = fUnquote(Request("vResc_Body"))
    vResc_Note          = fUnquote(Request("vResc_Note"))
  End Sub
  
  
  '...because we display the extracted values (above), we need to remove the double quotes
  Sub sRestoreResc
    vResc_ToName        = Replace(vResc_ToName, "''", "'")
    vResc_ToEmail       = Replace(vResc_ToEmail, "''", "'")
    vResc_Intro         = Replace(vResc_Intro, "''", "'")
    vResc_FromName      = Replace(vResc_FromName, "''", "'")
    vResc_FromEmail     = Replace(vResc_FromEmail, "''", "'")
    vResc_Subject       = Replace(vResc_Subject, "''", "'")
    vResc_Body          = Replace(vResc_Body, "''", "'")
    vResc_Note          = Replace(vResc_Note, "''", "'")
  End Sub
  

  Sub sInsertResc
    vSql = "SET NOCOUNT ON " _
         & "SET ANSI_WARNINGS ON " _
         & "INSERT INTO Resc (Resc_AcctId, Resc_ToName, Resc_ToEmail, Resc_Intro, Resc_Mods, Resc_ResTNo, Resc_MembNo, Resc_FromName, Resc_FromEmail, Resc_Subject, Resc_Body, Resc_Note) " _
         & "VALUES ('" & svCustAcctId & "', '" & vResc_ToName & "', '" & vResc_ToEmail & "', '" & vResc_Intro & "', '" & vResc_Mods & "', " & vResc_ResTNo & ", " & svMembNo & ", '" & vResc_FromName & "', '" & vResc_FromEmail & "', '" & vResc_Subject & "', '" & vResc_Body & "', '" & vResc_Note & "') " _
         & "SELECT vFld=@@IDENTITY " _
         & "SET NOCOUNT OFF "
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    vResc_No = oRs("vFld")
    sCloseDb
  End Sub


  Sub sUpdateRescDateSent
    vSql = "UPDATE Resc SET " _
         & "Resc_DateSent      = '" & Now()                 & "'  " _
         & "WHERE Resc_No      =  " & vResc_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sUpdateRescDateAccessed
    vSql = "UPDATE Resc SET " _
         & "Resc_LastDate      = '" & Now()                 & "', " _
         & "Resc_NoVisits      =      Resc_NoVisits + 1           " _
         & "WHERE Resc_No      =  " & vResc_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sUpdateResc
    vSql = "UPDATE Resc SET " _
         & "Resc_ResTNo        = '" & vResc_ResTNo          & "', " _
         & "Resc_ToName        = '" & vResc_ToName          & "', " _
         & "Resc_ToEmail       = '" & vResc_ToEmail         & "', " _
         & "Resc_Intro         = '" & vResc_Intro           & "', " _
         & "Resc_FromName      = '" & vResc_FromName        & "', " _
         & "Resc_FromEmail     = '" & vResc_FromEmail       & "', " _
         & "Resc_Subject       = '" & vResc_Subject         & "', " _
         & "Resc_Body          = '" & vResc_Body            & "', " _
         & "Resc_Note          = '" & vResc_Note            & "'  " _
         & "WHERE Resc_No      =  " & vResc_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sReadRecs_Rs
    vSql =  " SELECT * From FROM Resc WHERE Resc_Id = '" & vRescNo & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb 
  End Sub


  Sub sReadRecs_Rs
    vSql =  "SELECT Resc_Mods FROM Resc WHERE Resc_ToEmail = '" & vResc_ToEmail & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vResc_Mods = vResc_Mods & "," & oRs("Resc_Mods")
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb 
  End Sub


  '...Delete Resc
  Sub sDeleteResc (vRescNo)
    sOpenDb
    vSql = "DELETE FROM Resc WHERE Resc_No = " & vRescNo
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub
  

  '...Activate Resc
  Sub sActivateResc (vRescNo)
    sOpenDb
    vSql = "UPDATE Resc SET Resc_Active = 1 WHERE Resc_No = " & vRescNo
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  '...InActivate Resc
  Sub sInActivateResc (vRescNo)
    sOpenDb
    vSql = "UPDATE Resc SET Resc_Active = 0 WHERE Resc_No = " & vRescNo
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub



  Function fRescFrom (vRescRecTNo)
    Dim vSelected
    fRescFrom = ""
    vSql =  " SELECT DISTINCT Resc_RecTNo FROM Resc ORDER BY Resc_RecTNo"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      If oRs("Resc_RecTNo") = vRescRecTNo Then
        vSelected = " selected" 
      Else
        vSelected = ""
      End If
      fRescFrom = fRescFrom & "<option value=" & Chr(34) & oRs("Resc_RecTNo") & Chr(34) & vSelected & ">" & oRs("Resc_RecTNo") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb 
  End Function


  Function fRescMods (vRescToEmail)
    fRescMods = ""
    If Len(vRescToEmail) > 3 Then 
      vSql =  " SELECT Resc_Mods FROM Resc WHERE Resc_ToEmail = '" & vRescToEmail & "' ORDER BY Resc_Date DESC"
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        fRescMods = fRescMods & oRs("Resc_Mods") & ", "    
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
      If Len(fRescMods) > 6 Then 
        fRescMods = Left(fRescMods, Len(fRescMods) - 2)
      End If
    End If
  End Function


  Sub sGetRescByNo (vRescNo)
    vResc_Eof = True
    vSql = "SELECT * FROM Resc WHERE Resc_No = " & vRescNo & " AND Resc_Active = 1"
'   vSql = "SELECT * FROM Resc WHERE Resc_No = " & vRescNo
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadResc
      vResc_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub  


%>