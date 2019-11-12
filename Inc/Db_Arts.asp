<%
  Dim vArts_No, vArts_Type, vArts_Title, vArts_Keywords, vArts_Desc, vArts_Author, vArts_Article
  Dim vArts_Eof

  '...Get Arts Record
  Sub sGetArts
    If IsNumeric(vArts_No) Then 
      vArts_Eof = False
      vSql = "SELECT * FROM Arts WHERE Arts_No= " & vArts_No
      sOpenDb    
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then 
        sReadArts
        vArts_Eof = True
      End If
      Set oRs = Nothing
      sCloseDb
    End If
  End Sub

  '...Get Arts Title
  Function fArtsTitle
    fArtsTitle = ""
    If IsNumeric(vArts_No) Then 
      Dim oRs
      vSql = "SELECT Arts_Title FROM Arts WHERE Arts_No= " & vArts_No
      sOpenDb    
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then 
        fArtsTitle = oRs("Arts_Title")
      End If
      Set oRs = Nothing
      sCloseDb    
    End If
  End Function


  '...Get Arts Article
  Function fArtsArticle (vArts_No)
    fArtsArticle = ""
    If IsNumeric(vArts_No) Then 
      Dim oRs
      vSql = "SELECT Arts_Article FROM Arts WHERE Arts_No= " & vArts_No
      sOpenDb    
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then 
        fArtsArticle= oRs("Arts_Article")
      End If
      Set oRs = Nothing
      sCloseDb    
    Else
      fArtsArticle = "No Article Requested"
    End If
  End Function

  '...Get Arts Recordset
  Sub sGetArts_rs
    vSql = "SELECT * FROM Arts "
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
  End Sub

  Sub sReadArts
    vArts_No             = oRs("Arts_No")
    vArts_Type           = oRs("Arts_Type")
    vArts_Title          = oRs("Arts_Title")
    vArts_Keywords       = oRs("Arts_Keywords")
    vArts_Desc           = oRs("Arts_Desc")
    vArts_Author         = oRs("Arts_Author")
    vArts_Article        = oRs("Arts_Article") 
  End Sub

  Sub sExtractArts
    vArts_No             = Request.Form("vArts_No")
    vArts_Type           = fDefault(Request.Form("vArts_Type"),"A")
    vArts_Title          = fUnQuote(Request.Form("vArts_Title"))
    vArts_Keywords       = fUnQuote(Request.Form("vArts_Keywords"))
    vArts_Desc           = fUnQuote(Request.Form("vArts_Desc"))
    vArts_Author         = fUnQuote(Request.Form("vArts_Author"))
    vArts_Article        = fUnQuote(Request.Form("vArts_Article"))
  End Sub
  
  Sub sInsertArts
    vSql = "INSERT INTO Arts "
    vSql = vSql & "(Arts_Type, Arts_Title, Arts_Keywords, Arts_Desc, Arts_Author, Arts_Article)"
    vSql = vSql & " VALUES ('" & vArts_Type & "', '" & vArts_Title & "', '" & vArts_Keywords & "', '" & vArts_Desc & "', '" & vArts_Author & "', '" & vArts_Article & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sUpdateArts
    vSql = "UPDATE Arts SET"
    vSql = vSql & " Arts_Type     = '" & vArts_Type              & "', " 
    vSql = vSql & " Arts_Title    = '" & vArts_Title             & "', " 
    vSql = vSql & " Arts_Keywords = '" & vArts_Keywords          & "', " 
    vSql = vSql & " Arts_Desc     = '" & vArts_Desc              & "', " 
    vSql = vSql & " Arts_Author   = '" & vArts_Author            & "', "
    vSql = vSql & " Arts_Article  = '" & vArts_Article           & "'  "
    vSql = vSql & " WHERE Arts_No =  " & vArts_No
    sOpenDb
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub
  
  Sub sDeleteArts
    vSql = "DELETE FROM Arts WHERE Arts_No = " & vArts_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...get all Arts
  Function fArtsOptions
    Dim oRs
    fArtsOptions = ""
    sOpenDb
    vSql = "Select * FROM Arts"
    Set oRs = oDb.Execute(vSql)    
    Do While Not oRs.EOF 
      fArtsOptions = fArtsOptions & "<option>" & oRs("Arts_No") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop      
    sCloseDb           
  End Function
  

%>