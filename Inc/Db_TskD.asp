<%
  Dim vTskD_Key, vTskD_No, vTskD_Order, vTskD_Type, vTskD_Id, vTskD_Title, vTskD_Active, vTskD_Window
  Dim vTskD_Eof

  '...Get TskD Recordset
  Sub sGetTskD_Rs (vTskDNo)
    vSql = "SELECT * FROM TskD WHERE TskD_No = " & vTskDNo & " ORDER BY TskD_Order"
'   sDebug
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
  End Sub

  Function fTskD_ActiveRs (vTskDNo)
    fTskD_ActiveRs = False
    vSql = "SELECT * FROM TskD WHERE TskD_No = " & vTskDNo & " AND TskD_Active = 1"
'   sDebug
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fTskD_ActiveRs = True
    Set oRs2 = Nothing
    sCloseDb2      
  End Function

  Sub sGetTskD (vTskDKey)
    vTskD_Eof = False
    vSql = "SELECT * FROM TskD WHERE TskD_Key = " & vTskDKey
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      sReadTskD
      vTskD_Eof = True
    End If
    Set oRs2 = Nothing
    sCloseDb2    
  End Sub

  Sub sReadTskD
    vTskD_Key             = oRs2("TskD_Key")
    vTskD_No              = oRs2("TskD_No")
    vTskD_Order           = oRs2("TskD_Order")
    vTskD_Type            = oRs2("TskD_Type")
    vTskD_Id              = oRs2("TskD_Id")
    vTskD_Title           = oRs2("TskD_Title")
    vTskD_Active          = oRs2("TskD_Active")
    vTskD_Window          = oRs2("TskD_Window")
  End Sub

  Sub sExtractTskD
    vTskD_Key             = Request.Form("vTskD_Key")
    vTskD_No              = Request.Form("vTskD_No")
    vTskD_Order           = Request.Form("vTskD_Order")
    vTskD_Type            = Request.Form("vTskD_Type")
    vTskD_Id              = fUnquote(Request.Form("vTskD_Id"))    '...allow single quotes when adding in special vSql parms
    vTskD_Title           = fUnquote(Request.Form("vTskD_Title"))
    vTskD_Active          = fDefault(Request.Form("vTskD_Active"), 0)
    vTskD_Window          = fDefault(Request.Form("vTskD_Window"), 1)
    
    vOrderNo              = Cint(Request.Form("vOrderNo"))

  End Sub
  
  Sub sInsertTskD
    vSql = "INSERT INTO TskD "
    vSql = vSql & "(TskD_No, TskD_Order, TskD_Type, TskD_Id, TskD_Title, TskD_Active, TskD_Window)"
    vSql = vSql & " VALUES (" & vTskD_No & ", " & vTskD_Order & ", '" & vTskD_Type & "', '" & vTskD_Id & "', '" & fUnquote(vTskD_Title) & "', " & fSqlBoolean (vTskD_Active) & ", " & vTskD_Window & ")"
'   sDebug
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub

  Sub sInsertTskDEmpty
    '...get next order number so insert appears at end of the list
    vSql = "SELECT MAX(TskD_Order) + 1 AS TskD_Order FROM TskD WHERE (TskD_No = " & vTskD_No & ") GROUP BY TskD_No"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      vTskD_Order = oRs2("TskD_Order")
    Else
      vTskD_Order = 1
    End If    

    vSql = "INSERT INTO TskD (TskD_No, TskD_Order) VALUES (" & vTskD_No & ", " & vTskD_Order & ")"
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub

  '...insert cloned task assets
  Sub sInsertNextTskD (vTskDNo, vNextNo)
    vSql = "SELECT * FROM TskD WHERE TskD_No = " & vTskDNo & " ORDER BY TskD_Order"
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    Do While Not oRs2.Eof 
      sReadTskD
      vTskD_No = vNextNo
      vSql = "INSERT INTO TskD "
      vSql = vSql & "(TskD_No, TskD_Order, TskD_Type, TskD_Id, TskD_Title, TskD_Active, TskD_Window)"
      vSql = vSql & " VALUES (" & vTskD_No & ", " & vTskD_Order & ", '" & vTskD_Type & "', '" & vTskD_Id & "', '" & fUnquote(vTskD_Title) & "', " & fSqlBoolean (vTskD_Active) & ", " & vTskD_Window & ")"
'     sDebug
      oDb2.Execute(vSql)
      oRs2.MoveNext
    Loop
    Set oRs2 = Nothing
    sCloseDb2 
  End Sub

  Sub sUpdateTskD
    vSql = "UPDATE TskD SET"
    vSql = vSql & " TskD_No              =  " & vTskD_No                        & " , " 
    vSql = vSql & " TskD_Order           =  " & vTskD_Order                     & " , " 
    vSql = vSql & " TskD_Type            = '" & vTskD_Type                      & "', " 
    vSql = vSql & " TskD_Id              = '" & vTskD_Id                        & "', " 
    vSql = vSql & " TskD_Title           = '" & vTskD_Title                     & "', " 
    vSql = vSql & " TskD_Active          =  " & vTskD_Active                    & " , " 
    vSql = vSql & " TskD_Window          =  " & vTskD_Window
    vSql = vSql & " WHERE TskD_Key       =  " & vTskD_Key
    sOpenDb2
'   sDebug
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub
  
  Sub sUpdateTskDOrder
    vSql = "UPDATE TskD SET TskD_Order =  " & vTskD_Order & " WHERE TskD_Key =  " & vTskD_Key
    sOpenDb2
'   sDebug
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub
    
  Sub sDeleteTskD
    vSql = "DELETE FROM TskD WHERE TskD_Key = " & vTskD_Key
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub


%>