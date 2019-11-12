<%
  Dim vGrou_Id, vGrou_Active, vGrou_Title, vGrou_Desc, vGrou_Requires, vGrou_Supplier, vGrou_Programs
  Dim vGrou_Eof

  '...Get Grou RecordSet
  Sub sGetGrou_Rs
    vSql = "SELECT * FROM Grou WHERE (Grou_Active = 1) ORDER BY SUBSTRING(Grou_Id, 6, 2), Grou_Id"
    sOpenDbBase2
    Set oRsBase2 = oDbBase2.Execute(vSql)
  End Sub

  '...Get Grou Record
  Sub sGetGrou (vGrouId)
    vGrou_Eof = False
    vSql = "SELECT * FROM Grou WHERE Grou_Id= '" & vGrouId & "'"
    sOpenDbBase2    
    Set oRsBase2 = oDbBase2.Execute(vSql)
    If Not oRsBase2.Eof Then 
      sReadGrou
      vGrou_Eof = True
    End If
    Set oRsBase2 = Nothing
    sCloseDbBase2    
  End Sub

  Sub sReadGrou
    vGrou_Id             = oRsBase2("Grou_Id")
    vGrou_Active         = oRsBase2("Grou_Active")
    vGrou_Title          = oRsBase2("Grou_Title")
    vGrou_Desc           = oRsBase2("Grou_Desc")
    vGrou_Requires       = oRsBase2("Grou_Requires") 
    vGrou_Supplier       = oRsBase2("Grou_Supplier")
    vGrou_Programs       = oRsBase2("Grou_Programs")
  End Sub

  Sub sExtractGrou
    vGrou_Id             = Ucase(Request.Form("vGrou_Id"))
    vGrou_Active         = Request.Form("vGrou_Active")
    vGrou_Title          = Request.Form("vGrou_Title")
    vGrou_Desc           = Request.Form("vGrou_Desc")
    vGrou_Requires       = Request.Form("vGrou_Requires")
    vGrou_Supplier       = Request.Form("vGrou_Supplier")
    vGrou_Programs       = fNoQuote(Request.Form("vGrou_Programs"))
  End Sub
  
  Sub sInsertGrou
    vSql = "INSERT INTO Grou "
    vSQL = vSQL & "(Grou_Id, Grou_Active, Grou_Title, Grou_Desc, Grou_Requires, Grou_Supplier, Grou_Programs)"
    vSQL = vSQL & " VALUES ('" & vGrou_Id & "', " & vGrou_Active & ", '" & fUnquote(vGrou_Title) & "', '" & fUnquote(vGrou_Desc) & "', '" & fUnquote(vGrou_Requires) & "', '" & fUnquote(vGrou_Supplier) & "', '" & vGrou_Programs & "')"
'   sDebug
    sOpenDbBase2
    oDbBase2.Execute(vSQL)
    sCloseDbBase2
  End Sub

  Sub sUpdateGrou
    vSQL = "UPDATE Grou SET"
    vSQL = vSQL & " Grou_Active   =  " & fUnquote(vGrou_Active)   & " , " 
    vSQL = vSQL & " Grou_Title    = '" & fUnquote(vGrou_Title)    & "', " 
    vSQL = vSQL & " Grou_Desc     = '" & fUnquote(vGrou_Desc)     & "', " 
    vSQL = vSQL & " Grou_Requires = '" & fUnquote(vGrou_Requires) & "', " 
    vSQL = vSQL & " Grou_Supplier = '" & fUnquote(vGrou_Supplier) & "', " 
    vSQL = vSQL & " Grou_Programs = '" & vGrou_Programs           & "'  " 
    vSQL = vSQL & " WHERE Grou_Id = '" & vGrou_Id                 & "'  "
    sOpenDbBase2
    oDbBase2.Execute(vSQL)
    sCloseDbBase2
  End Sub

  
  Sub sDeleteGrou
    vSQL = "DELETE FROM Grou WHERE Grou_Id = '" & vGrou_Id & "'"
    sOpenDbBase2
    oDbBase2.Execute(vSQL)
    sCloseDbBase2
  End Sub


  '...return group subsets with program strings for channel setup
  Function fGrouDropdown (vIds, vSubSet)
    Dim vCurrentId, vSelected, aProg, vFree, vEcom
    '...save the current Group id
    fGrouDropDown = vbCrLf
    vSql = "SELECT * FROM Grou WHERE (Grou_Active =1) AND Len(Grou_Programs) > 0 AND Right(Grou_Id, 2) = '" & fIf(Ucase(svLang) = "ROOT", "EN", svLang)  & "' AND Len(Grou_Title) > 0" 
    If Len(vSubSet) > 0 Then
    vSql = vSql & " AND Grou_Subset = '" & vSubSet & "'"
    End If
'   sDebug
    sOpenDbBase2
    Set oRsBase2 = oDbBase2.Execute(vSql)
    Do While Not oRsBase2.Eof
      sReadGrou
      '...If group includes items for sell, add $ after the group name
      aProg = Split(vGrou_Programs)
      vEcom = "" : vFree = ""
      For i = 0 To Ubound(aProg)
        If Mid(aProg(i), 8, 3) = "~0~" Then
          vFree = "F"
        Else
          vEcom = "$"
        End If
      Next 

      If Instr(vIds, vGrou_Id) > 0 Then
        vSelected = " selected" 
      Else
        vSelected = ""
      End If
      i = "          <option value=" & Chr(34) & vGrou_Id & Chr(34) & vSelected & ">" & fLeft(vGrou_Title, 48) & " [" & vFree & vEcom & "]</option>" & vbCrLf
      fGrouDropdown = fGrouDropdown & i
      oRsBase2.MoveNext	        
    Loop
    Set oRsBase2 = Nothing      
    sCloseDbBase2
    '...save the current Grouaign id
    vGrou_Id = vIds
  End Function 


  '...Get Grou Title
  Function fGrouTitle (vGrouId)
    Dim oRs2
    fGrouTitle = ""
    vSql = "SELECT Grou_Title FROM Grou WHERE Grou_Id= '" & vGrouId & "'"
    sOpenDbBase2    
    Set oRs2 = oDbBase2.Execute(vSql)
    If Not oRs2.Eof Then 
      fGrouTitle = oRs2("Grou_Title")
    End If
    Set oRs2 = Nothing
    sCloseDbBase2
  End Function
%>