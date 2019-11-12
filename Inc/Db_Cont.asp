<%
  '____ Cont  ________________________________________________________________________

  Dim vCont_Id, vCont_Title, vCont_Value1, vCont_Value2, vCont_Value3, vCont_Value4, vCont_Value5, vCont_Value6, vCont_Desc1, vCont_Desc2, vCont_Desc3, vCont_Desc4, vCont_Desc5, vCont_Desc6
  Dim vCont_Eof

 
  '...Get Contomer RecordSet
  Sub sGetCont_rs   
    vSql = "SELECT * FROM Cont"
    sOpenDB
    Set oRs = oDB.Execute(vSql)
  End Sub  

  '...Get Cont Recordset
  Sub sGetCont (vContId)
    vSql = "SELECT * FROM Cont WHERE Cont_Id= '" & vContId & "'"
    sOpenDB    
    Set oRs = oDB.Execute(vSql)
    If Not oRs.Eof Then 
      sReadCont
      vCont_Eof = False
    Else
      vCont_Eof = True
    End If
    Set oRs = Nothing
    sCloseDB    
  End Sub

  Sub sReadCont
    vCont_Id                = oRs("Cont_Id")
    vCont_Title             = oRs("Cont_Title")
    vCont_Value1            = oRs("Cont_Value1")
    vCont_Value2            = oRs("Cont_Value2")
    vCont_Value3            = oRs("Cont_Value3")
    vCont_Value4            = oRs("Cont_Value4")
    vCont_Value5            = oRs("Cont_Value5")
    vCont_Value6            = oRs("Cont_Value6")
    vCont_Desc1             = oRs("Cont_Desc1")
    vCont_Desc2             = oRs("Cont_Desc2")
    vCont_Desc3             = oRs("Cont_Desc3")
    vCont_Desc4             = oRs("Cont_Desc4")
    vCont_Desc5             = oRs("Cont_Desc5")
    vCont_Desc6             = oRs("Cont_Desc6")
  End Sub

  Sub sExtractCont
    vCont_Id                = Request.Form("vCont_Id")
    vCont_Title             = fUnquote(Request.Form("vCont_Title"))
    vCont_Value1            = fDefault(Request.Form("vCont_Value1"), 0)
    vCont_Value2            = fDefault(Request.Form("vCont_Value2"), 0)
    vCont_Value3            = fDefault(Request.Form("vCont_Value3"), 0)
    vCont_Value4            = fDefault(Request.Form("vCont_Value4"), 0)
    vCont_Value5            = fDefault(Request.Form("vCont_Value5"), 0)
    vCont_Value6            = fDefault(Request.Form("vCont_Value6"), 0)
    vCont_Desc1             = fUnquote(Request.Form("vCont_Desc1"))
    vCont_Desc2             = fUnquote(Request.Form("vCont_Desc2"))
    vCont_Desc3             = fUnquote(Request.Form("vCont_Desc3"))
    vCont_Desc4             = fUnquote(Request.Form("vCont_Desc4"))
    vCont_Desc5             = fUnquote(Request.Form("vCont_Desc5"))
    vCont_Desc6             = fUnquote(Request.Form("vCont_Desc6"))
  End Sub
  
  Sub sUpdateCont
    sOpenDb
    On Error Resume Next
    '...try to insert new record
    vSql = "SET ANSI_WARNINGS ON "
    vSql = vSql & "INSERT INTO Cont "
    vSql = vSql & "(Cont_Id, Cont_Title, Cont_Value1, Cont_Value2, Cont_Value3, Cont_Value4, Cont_Value5, Cont_Value6, Cont_Desc1, Cont_Desc2, Cont_Desc3, Cont_Desc4, Cont_Desc5, Cont_Desc6)"
    vSql = vSql & " VALUES ('" & vCont_Id & "', '" & vCont_Title & "', " & vCont_Value1 & ", " & vCont_Value2 & ", " & vCont_Value3 & ", " & vCont_Value4 & ", " & vCont_Value5 & ", " & vCont_Value6 & ", '" & vCont_Desc1 & "', '" & vCont_Desc2 & "', '" & vCont_Desc3 & "', '" & vCont_Desc4 & "', '" & vCont_Desc5 & "', ' " & vCont_Desc6 & "')"
 '  sDebug
    oDb.Execute(vSql)

    '...if already on file then update
    If Err.Number <> 0 Or Err.Number <> "" Then 
      vSql = "SET ANSI_WARNINGS ON "
      vSql = vSql & "UPDATE Cont SET"
      vSql = vSql & " Cont_Title             = '" & vCont_Title              & "', " 
      vSql = vSql & " Cont_Value1            =  " & vCont_Value1             & " , " 
      vSql = vSql & " Cont_Value2            =  " & vCont_Value2             & " , " 
      vSql = vSql & " Cont_Value3            =  " & vCont_Value3             & " , " 
      vSql = vSql & " Cont_Value4            =  " & vCont_Value4             & " , " 
      vSql = vSql & " Cont_Value5            =  " & vCont_Value5             & " , " 
      vSql = vSql & " Cont_Value6            =  " & vCont_Value6             & " , " 
      vSql = vSql & " Cont_Desc1             = '" & vCont_Desc1              & "', " 
      vSql = vSql & " Cont_Desc2             = '" & vCont_Desc2              & "', " 
      vSql = vSql & " Cont_Desc3             = '" & vCont_Desc3              & "', " 
      vSql = vSql & " Cont_Desc4             = '" & vCont_Desc4              & "', " 
      vSql = vSql & " Cont_Desc5             = '" & vCont_Desc5              & "', " 
      vSql = vSql & " Cont_Desc6             = '" & vCont_Desc6              & "'  " 
      vSql = vSql & " WHERE Cont_Id          = '" & vCont_Id                 & "'  "
 '    sDebug
      oDb.Execute(vSql)
      sCloseDb
    End If    
    sCloseDb
  End Sub

  Sub sDeleteCont
    vSql = "DELETE FROM Cont WHERE Cont_Id = '" & vCont_Id & "'"
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...Get Contract RecordSet
  Function fContDropDown (vContId)
    fContDropDown = vbCrLf
    sGetCont_rs
    Do While Not oRs.Eof
      sReadCont
      fContDropDown = fContDropDown & "<option value='" & vCont_Id & "' " & fSelect(vContId, vCont_Id) & ">" & vCont_Id & "&nbsp;&nbsp;&nbsp;&nbsp;" & Replace(fLeft(vCont_Title, 32), vbCrLf, " ") & "</option>" & vbCrLf
    oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDB    
  End Function

%>