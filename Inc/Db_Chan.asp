<%
  Dim vChan_Id, vChan_Title
  Dim vChan_2004e, vChan_2005e, vChan_2006e, vChan_2007e, vChan_2008e, vChan_2009e, vChan_2010e, vChan_2011e, vChan_2012e
  Dim vChan_2004m, vChan_2005m, vChan_2006m, vChan_2007m, vChan_2008m, vChan_2009m, vChan_2010m, vChan_2011m, vChan_2012m
  Dim vChan_Owner, vChan_Contacts, vChan_Notes
  Dim bChan_Eof


  '...Get Chan Recordset
  Sub sGetChan_Rs
    vSql = "SELECT * FROM Chan"
'   sDebug
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
  End Sub


  Sub sGetChan
    bChan_Eof = True
    vSql = "SELECT * FROM Chan WHERE Chan_Id = '" & vChan_Id & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadChan
      bChan_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb    
  End Sub


  Sub sReadChan
    vChan_Id       = oRs("Chan_Id")
    vChan_Title    = oRs("Chan_Title")
    vChan_2004e     = oRs("Chan_2004e")
    vChan_2005e     = oRs("Chan_2005e")
    vChan_2006e     = oRs("Chan_2006e")
    vChan_2007e     = oRs("Chan_2007e")
    vChan_2008e     = oRs("Chan_2008e")
    vChan_2009e     = oRs("Chan_2009e")
    vChan_2010e     = oRs("Chan_2010e")
    vChan_2011e     = oRs("Chan_2011e")
    vChan_2012e     = oRs("Chan_2012e")

    vChan_2004m     = oRs("Chan_2004m")
    vChan_2005m     = oRs("Chan_2005m")
    vChan_2006m     = oRs("Chan_2006m")
    vChan_2007m     = oRs("Chan_2007m")
    vChan_2008m     = oRs("Chan_2008m")
    vChan_2009m     = oRs("Chan_2009m")
    vChan_2010m     = oRs("Chan_2010m")
    vChan_2011m     = oRs("Chan_2011m")
    vChan_2012m     = oRs("Chan_2012m")

    vChan_Owner    = oRs("Chan_Owner")
    vChan_Contacts = oRs("Chan_Contacts")
    vChan_Notes    = oRs("Chan_Notes")
  End Sub


  Sub sExtractChan
    vChan_Id       = Request.Form("vChan_Id")
    vChan_Title    = fUnquote(Request.Form("vChan_Title"))
    vChan_Owner    = fUnquote(Request.Form("vChan_Owner"))   
    vChan_Contacts = fUnquote(Request.Form("vChan_Contacts"))   
    vChan_Notes    = fUnquote(Request.Form("vChan_Notes"))   
  End Sub
  

  Sub sUpdateChan
    vSql = "UPDATE Chan SET"
    vSql = vSql & " Chan_Title             = '" & vChan_Title      & "', " 
    vSql = vSql & " Chan_Owner             = '" & vChan_Owner      & "', " 
    vSql = vSql & " Chan_Contacts          = '" & vChan_Contacts   & "', " 
    vSql = vSql & " Chan_Notes             = '" & vChan_Notes      & "'  " 
    vSql = vSql & " WHERE Chan_Id          = '" & vChan_Id         & "' "
    sOpenDb
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  
  Sub sDeleteChan
    vSql = "DELETE FROM Chan WHERE Chan_Id = '" & vChan_Id & "'"
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub



%>