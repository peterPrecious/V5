<%
  Dim vClus_Id, vClus_Tab1, vClus_Tab2, vClus_Tab3, vClus_Tab4, vClus_Tab5, vClus_Tab6, vClus_Tab1_Name, vClus_Tab2_Name, vClus_Tab3_Name, vClus_Tab4_Name, vClus_Tab5_Name, vClus_Tab6_Name, vClus_Tab7_Name, vClus_Tab8_Name
  Dim vClus_Eof


  '____ Clus ________________________________________________________________________

  Sub sGetClus
    vClus_Eof = False
    vSql = "SELECT * FROM Clus WHERE Clus_Id = '" & vClus_Id & "'"
'   sDebug
    sOpenDb
    On Error Resume Next
    Set oRs = oDb.Execute(vSql)    
    If Err.Number = 0 or Err.Number = "" Then 
      vClus_Eof = False
      sReadClus
    End If
    On Error GoTo 0          
    Set oRs = Nothing      
    sCloseDb
  End Sub      

  Sub sReadClus
    vClus_Id        = oRs("Clus_Id")
    vClus_Tab1      = oRs("Clus_Tab1")
    vClus_Tab2      = oRs("Clus_Tab2")
    vClus_Tab3      = oRs("Clus_Tab3")
    vClus_Tab4      = oRs("Clus_Tab4")
    vClus_Tab5      = oRs("Clus_Tab5")
    vClus_Tab6      = oRs("Clus_Tab6")
    vClus_Tab1_Name = oRs("Clus_Tab1_Name")
    vClus_Tab2_Name = oRs("Clus_Tab2_Name")
    vClus_Tab3_Name = oRs("Clus_Tab3_Name")
    vClus_Tab4_Name = oRs("Clus_Tab4_Name")
    vClus_Tab5_Name = oRs("Clus_Tab5_Name")
    vClus_Tab6_Name = oRs("Clus_Tab6_Name")
    vClus_Tab7_Name = oRs("Clus_Tab7_Name")
    vClus_Tab8_Name = oRs("Clus_Tab8_Name")
  End Sub

  '...insert record unless already on file, then update
  Sub sInsertClus
    vSQL = "INSERT INTO Clus"
    vSQL = vSQL & " (Clus_Id, Clus_Tab1, Clus_Tab2, Clus_Tab3, Clus_Tab4, Clus_Tab5, Clus_Tab6, Clus_Tab1_Name, Clus_Tab2_Name, Clus_Tab3_Name, Clus_Tab4_Name, Clus_Tab5_Name, Clus_Tab6_Name, Clus_Tab7_Name, Clus_Tab8_Name)"
    vSQL = vSQL & " VALUES ('" & vClus_Id & "', " & vClus_Tab1 & ", " & vClus_Tab2 & ", " & vClus_Tab3 & ", " & vClus_Tab4 & ", " & vClus_Tab5 & ", " & vClus_Tab6 & ", '" & vClus_Tab1_Name & "', '" & vClus_Tab2_Name & "', '" & vClus_Tab3_Name & "', '" & vClus_Tab4_Name & "', '" & vClus_Tab5_Name & "', '" & vClus_Tab6_Name & "', '" & vClus_Tab7_Name & "', '" & vClus_Tab8_Name & "')"
'   sDebug
    sOpenDb
    On Error Resume Next 
    Set oRs = oDb.Execute(vSql)
    If Err.Number <> 0 Or Err.Number <> "" Then 
      On Error GoTo 0          
      sCloseDb
      sUpdateClus
      Exit Sub
    End If
    On Error GoTo 0          
    Set oRs = Nothing      
    sCloseDb
  End Sub
  
  Sub sUpdateClus
    vSQL = "UPDATE Clus SET"
    vSQL = vSQL & " Clus_Tab1        =  " & vClus_Tab1       & " ,  " 
    vSQL = vSQL & " Clus_Tab2        =  " & vClus_Tab2       & " ,  " 
    vSQL = vSQL & " Clus_Tab3        =  " & vClus_Tab3       & " ,  " 
    vSQL = vSQL & " Clus_Tab4        =  " & vClus_Tab4       & " ,  " 
    vSQL = vSQL & " Clus_Tab5        =  " & vClus_Tab5       & " ,  " 
    vSQL = vSQL & " Clus_Tab6        =  " & vClus_Tab6       & " ,  " 
    vSQL = vSQL & " Clus_Tab1_Name   = '" & vClus_Tab1_Name  & "',  " 
    vSQL = vSQL & " Clus_Tab2_Name   = '" & vClus_Tab2_Name  & "',  " 
    vSQL = vSQL & " Clus_Tab3_Name   = '" & vClus_Tab3_Name  & "',  " 
    vSQL = vSQL & " Clus_Tab4_Name   = '" & vClus_Tab4_Name  & "',  " 
    vSQL = vSQL & " Clus_Tab5_Name   = '" & vClus_Tab5_Name  & "',  " 
    vSQL = vSQL & " Clus_Tab6_Name   = '" & vClus_Tab6_Name  & "',   " 
    vSQL = vSQL & " Clus_Tab7_Name   = '" & vClus_Tab7_Name  & "',  " 
    vSQL = vSQL & " Clus_Tab8_Name   = '" & vClus_Tab8_Name  & "'   " 
    vSQL = vSQL & " WHERE Clus_Id    = '" & vClus_Id         & "'   " 
    sOpenDb 
 '  sDebug
    oDb.Execute(vSQL)
    sCloseDb
  End Sub

  Sub sExtractClus  
    vClus_Id           = Request.Form("vClus_Id")
    vClus_Tab1         = Request.Form("vClus_Tab1")
    vClus_Tab2         = Request.Form("vClus_Tab2")
    vClus_Tab3         = Request.Form("vClus_Tab3")
    vClus_Tab4         = Request.Form("vClus_Tab4")
    vClus_Tab5         = Request.Form("vClus_Tab5")
    vClus_Tab6         = Request.Form("vClus_Tab6")
    vClus_Tab1_Name    = Request.Form("vClus_Tab1_Name")
    vClus_Tab2_Name    = Request.Form("vClus_Tab2_Name")
    vClus_Tab3_Name    = Request.Form("vClus_Tab3_Name")
    vClus_Tab4_Name    = Request.Form("vClus_Tab4_Name")
    vClus_Tab5_Name    = Request.Form("vClus_Tab5_Name")
    vClus_Tab6_Name    = Request.Form("vClus_Tab6_Name")
    vClus_Tab7_Name    = Request.Form("vClus_Tab7_Name")
    vClus_Tab8_Name    = Request.Form("vClus_Tab8_Name")

    If fNoValue(vClus_Tab1) Then vClus_Tab1 = 0
    If fNoValue(vClus_Tab2) Then vClus_Tab2 = 0
    If fNoValue(vClus_Tab3) Then vClus_Tab3 = 0
    If fNoValue(vClus_Tab4) Then vClus_Tab4 = 0
    If fNoValue(vClus_Tab5) Then vClus_Tab5 = 0
    If fNoValue(vClus_Tab6) Then vClus_Tab6 = 0
  End Sub

  Sub sDeleteClus
    vSQL = "DELETE FROM Clus WHERE Clus_Id = '" & vClus_Id & "'"
    sOpenDb
    oDb.Execute(vSQL)
    sCloseDb
  End Sub
  
  Function fClusDropdown (vId) 
    Dim vCurrentId, vSelected
    '...save the current Cluster Id
    vCurrentId = vId
    fClusDropDown = vbCrLf
    vSql = "SELECT Clus_Id FROM Clus"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vClus_Id = oRs("Clus_Id")
      If vClus_Id = vCurrentId Then
        vSelected = " selected" 
      Else
        vSelected = ""
      End If
      i = "          <option value=" & Chr(34) & vClus_Id & Chr(34) & vSelected & ">" & vClus_Id & "</option>" & vbCrLf
      fClusDropdown = fClusDropdown & i
      oRs.MoveNext	        
    Loop
    Set oRs = Nothing      
    sCloseDb
    '...save the current Clusaign id
    vClus_Id = vCurrentId
  End Function   
%>