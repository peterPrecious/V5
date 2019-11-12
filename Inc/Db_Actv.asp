<%
  Dim vActv_Id, vActv_No, vActv_AccessOk, vActv_AccessNo, vActv_Title, vActv_Desc, vActv_Active, vActv_Length, vActv_AlteredOn, vActv_AlteredBy
  Dim vActv_Eof
  
  '____ Actv ________________________________________________________________________

  '...Get Actv Recordset
  Sub sGetActv (vActvId)
    vSql = "SELECT * FROM Actv WHERE Actv_Id= '" & vActvId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadActv
      vActv_Eof = False
    Else
      vActv_Eof = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Sub



  '...Get Actv Record by No
  Sub sGetActvByNo (vActvNo)
    vSql = "SELECT * FROM Actv WHERE Actv_No= " & vActvNo
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadActv
      vActv_Eof = False
    Else
      vActv_Eof = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Sub



  Sub sReadActv
    vActv_Id               = Ucase(oRsBase("Actv_Id"))
    vActv_No               = oRsBase("Actv_No")
    vActv_AccessOk         = oRsBase("Actv_AccessOk")
    vActv_AccessNo         = oRsBase("Actv_AccessNo")
    vActv_Title            = oRsBase("Actv_Title")
    vActv_Desc             = oRsBase("Actv_Desc")
    vActv_Active           = oRsBase("Actv_Active")
    vActv_Length           = oRsBase("Actv_Length")
    vActv_AlteredOn				 = oRsBase("Actv_AlteredOn")
    vActv_AlteredBy 			 = oRsBase("Actv_AlteredBy")
  End Sub

  
  Sub sExtractActv
    vActv_Id               = Ucase(Request.Form("vActv_Id"))
    vActv_AccessOk         = Request.Form("vActv_AccessOk")
    vActv_AccessNo         = Request.Form("vActv_AccessNo")
    vActv_Title            = Request.Form("vActv_Title")
    vActv_Active           = Request.Form("vActv_Active")
    vActv_Desc             = Request.Form("vActv_Desc")
    vActv_Length           = Request.Form("vActv_Length")
    vActv_AlteredOn				 = Request.Form("vActv_AlteredOn")
    vActv_AlteredBy				 = Request.Form("vActv_AlteredBy")
  End Sub
  

%>