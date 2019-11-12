<%
  '____ Cust  ________________________________________________________________________

  Dim vCust_Id, vCust_AcctId, vCust_Title, vCust_Lang, vCust_Catalogue, vCust_Banner, vCust_Url
  Dim vCust_Level, vCust_Tab2, vCust_Tab3, vCust_Tab5

  Dim vCust_Eof

  '...Get Customer RecordSet 
  Sub sGetCust_Rs   
    vSql  = "SELECT * FROM Cust"
    sOpenDB
    Set oRs  = oDB.Execute(vSql)
  End Sub  

  Sub sGetCust_Rs_AcctId
    vSql  = "SELECT * FROM Cust ORDER BY Cust_AcctId"
    sOpenDB
    Set oRs  = oDB.Execute(vSql)
  End Sub  

  '...Get Cust Recordset
  Sub sGetCust (vCustId)
    vCust_Eof  = True
    If Len(Session("HostDb"))  = 0 Then Response.Redirect "Timeout.asp?vPage=" & Request.ServerVariables("Path_Info")
    vSql  = "SELECT * FROM Cust WHERE Cust_Id = '" & vCustId & "'"
    sOpenDb    
    Set oRs  = oDB.Execute(vSql)
    If Not oRs.Eof Then 
      sReadCust
      vCust_Eof  = False
    End If
    Set oRs  = Nothing
    sCloseDB    
  End Sub

  Sub sReadCust
    vCust_Id                  = oRs("Cust_Id")
    vCust_AcctId              = oRs("Cust_AcctId")
    vCust_Title               = oRs("Cust_Title")
    vCust_Lang                = oRs("Cust_Lang")
    vCust_Catalogue           = oRs("Cust_Groups")
    vCust_Banner              = oRs("Cust_Banner")
    vCust_Url                 = oRs("Cust_Url")
    vCust_FacilitatorId       = oRs("Cust_FacilitatorId")
    vCust_ManagerId           = oRs("Cust_ManagerId")
  End Sub

  Sub sExtractCust
    vCust_Id                  = Request.Form("vCust_Id")
    vCust_AcctId              = Request.Form("vCust_AcctId")
    vCust_Title               = Request.Form("vCust_Title")
    vCust_Lang                = fDefault(Request.Form("vCust_Lang"), "EN")
    vCust_Catalogue           = Trim(Replace(Request.Form("vCust_Catalogue"), ",", ""))
    vCust_Banner              = fDefault(Request.Form("vCust_Banner"), "vubz.jpg")
    vCust_Url                 = fDefault(Request.Form("vCust_Url"), "vubiz.com") 
    vCust_FacilitatorId       = Trim(fNoQuote(fDefault(Request.Form("vCust_FacilitatorId"), vCust_Id & "_FAC"))) 
    vCust_ManagerId           = Trim(fNoQuote(fDefault(Request.Form("vCust_ManagerId"), vCust_Id & "_MGR"))) 
  End Sub

  Sub sInsertChannel
    vSql  = vSql & "INSERT INTO Cust (Cust_Id, Cust_AcctId, Cust_Title, Cust_Lang, Cust_Groups, Cust_Banner, Cust_Url) "
    vSql  = vSql & "VALUES ('" & vCust_Id & "', '" & vCust_AcctId & "', '" & vCust_Title & "', '" & vCust_Lang & "', '" & vCust_Catalogue & "', '" & vCust_Banner & "', '" & vCust_Url & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb

    sAddInternalMemb vCust_AcctId

  End Sub

  '...same as above with a few different defaults (tabs, levels)
  Sub sInsertCorporate
    vSql  = vSql & "INSERT INTO Cust (Cust_Id, Cust_AcctId, Cust_Title, Cust_Lang, Cust_Groups, Cust_Banner, Cust_Url, Cust_Level, Cust_Tab2, Cust_Tab3, Cust_Tab5) "
    vSql  = vSql & "VALUES ('" & vCust_Id & "', '" & vCust_AcctId & "', '" & vCust_Title & "', '" & vCust_Lang & "', '" & vCust_Catalogue & "', '" & vCust_Banner & "', '" & vCust_Url & "', 4, 1, 0, 0)"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb

    sAddInternalMemb vCust_AcctId      
    If vCustLevel = 4 Then sSetupRepository vNewAcctId
  End Sub


  Sub sUpdateCust
    vSql  = vSql & "UPDATE Cust SET"
    vSql  = vSql & " Cust_Title              = '" & fUnquote(vCust_Title)    & "', " 
    vSql  = vSql & " Cust_Lang               = '" & vCust_Lang               & "', " 
    vSql  = vSql & " Cust_Groups             = '" & vCust_Catalogue          & "', " 
    vSql  = vSql & " Cust_Banner             = '" & vCust_Banner             & "', " 
    vSql  = vSql & " Cust_Url                = '" & vCust_Url                & "'  " 
    vSql  = vSql & " WHERE Cust_Id  = '" & vCust_Id                          & "'  "
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sDeleteCust
    sOpenDB    
    oDb.Execute("DELETE FROM Cust WHERE Cust_AcctId = '" & vCust_AcctId & "'")
    oDb.Execute("DELETE FROM Logs WHERE Logs_AcctId = '" & vCust_AcctId & "'")
    oDb.Execute("DELETE FROM Memb WHERE Memb_AcctId = '" & vCust_AcctId & "'")
    sCloseDB    
  End Sub

  Function fNextCustNo (vAcctId)
    fNextCustNo = 0
    vSql = "SELECT TOP 1 Cust_AcctId FROM Cust WHERE Cust_AcctId > '" & Left(vAcctId & "0000", 4) & "' AND Cust_AcctId < '" & Left(vAcctId & "9999", 4)  & "' ORDER BY Cust_AcctId DESC"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      fNextCustNo = Left(vAcctId & "0000", 4)
    Else
      fNextCustNo = oRs("Cust_AcctId")
    End If
    Set oRs = Nothing      
    sCloseDb
    fNextCustNo = fNextCustNo + 1
  End Function

%>