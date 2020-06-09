<%
  Dim vClss_Id, vClss_No, vClss_AccessOk, vClss_AccessNo, vClss_Title, vClss_Desc, vClss_Active, vClss_Length, vClss_AlteredOn, vClss_AlteredBy
  Dim vClss_Eof
  
  '____ Clss ________________________________________________________________________

  '...Get Clss Recordset
  Sub sGetClss (vClssId)
    vSql = "SELECT * FROM Clss WHERE Clss_Id= '" & vClssId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadClss
      vClss_Eof = False
    Else
      vClss_Eof = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Sub


  '...Get Clss Record by No
  Sub sGetClssByNo (vClssNo)
    vSql = "SELECT * FROM Clss WHERE Clss_No= " & vClssNo
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadClss
      vClss_Eof = False
    Else
      vClss_Eof = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Sub



  Sub sReadClss
    vClss_Id               = Ucase(oRsBase("Clss_Id"))
    vClss_No               = oRsBase("Clss_No")
    vClss_AccessOk         = oRsBase("Clss_AccessOk")
    vClss_AccessNo         = oRsBase("Clss_AccessNo")
    vClss_Title            = oRsBase("Clss_Title")
    vClss_Desc             = oRsBase("Clss_Desc")
    vClss_Active           = oRsBase("Clss_Active")
    vClss_Length           = oRsBase("Clss_Length")
    vClss_AlteredOn				 = oRsBase("Clss_AlteredOn")
    vClss_AlteredBy 			 = oRsBase("Clss_AlteredBy")
  End Sub

  
  Sub sExtractClss
    vClss_Id               = Ucase(Request.Form("vClss_Id"))
    vClss_AccessOk         = Request.Form("vClss_AccessOk")
    vClss_AccessNo         = Request.Form("vClss_AccessNo")
    vClss_Title            = Request.Form("vClss_Title")
    vClss_Active           = Request.Form("vClss_Active")
    vClss_Desc             = Request.Form("vClss_Desc")
    vClss_Length           = Request.Form("vClss_Length")
    vClss_AlteredOn				 = Request.Form("vClss_AlteredOn")
    vClss_AlteredBy				 = Request.Form("vClss_AlteredBy")
  End Sub
  


%>