<%
  '____ Elog - used to log transactions that go through EcomPatience.asp (mainly from Ecom2Checkout.asp and Ecom3Checout.asp  ___________________________________________

  Dim vElog_No, vElog_CustId, vElog_Id, vElog_Posted, vElog_Data
  Dim vElog_Eof


  Sub sReadElog
    vElog_No           = oRs("Elog_No")
    vElog_CustId       = oRs("Elog_CustId")
    vElog_Id           = oRs("Elog_Id")
    vElog_Posted       = oRs("Elog_Posted")
    vElog_Data         = oRs("Elog_Data")
  End Sub


  '...Get a list of all Elog for this account containing Elog_Posted and Elog_No
  Sub spElogById (vCustId)
    sOpenCmd
    With oCmd
      .CommandText = "spElogSelect"
      .Parameters.Append .CreateParameter("@Elog_CustId", adChar,     adParamInput,     8, vCustId)
      .Parameters.Append .CreateParameter("@Elog_CustId", adVarChar,  adParamInput,    64, vId)
    End With
    Set oRs = oCmd.Execute()
    If Not oRs.Eof Then 
      sReadElog
    Else
      sInitElog
    End If
    Set oRs = Nothing      
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...Get a specific Elog for this account containing Elog_Data
  Sub spElogByNo (vElogNo)
    sOpenCmd
    With oCmd
      .CommandText = "spElogSelect"
      .Parameters.Append .CreateParameter("@Elog_CustId", adInteger,  adParamInput,     8, vCustId)
    End With
    Set oRs = oCmd.Execute()
    If Not oRs.Eof Then 
      sReadElog
    Else
      sInitElog
    End If
    Set oRs = Nothing      
    Set oCmd = Nothing
    sCloseDb
  End Sub


  Sub spElogInsert (vCustId, vId, vData)  
    '...clean and truncate
    vData = Replace(vData, vbCrLf, "")
    vData = Replace(vData, vbTab, "")
    vData = Replace(vData, "  ", " ")
    vData = Replace(vData, "  ", " ")
    vData = Replace(vData, "  ", " ")
    vData = Replace(vData, "> <", "><")
    vData = Trim(vData)

    If Len(vData) > 8000 Then vData = Left(vData, 8000)    
    sOpenCmd
    With oCmd
      .CommandText = "spElogInsert"
      .Parameters.Append .CreateParameter("@Elog_CustId", adChar,     adParamInput,     8, vCustId)
      .Parameters.Append .CreateParameter("@Elog_Id",     adVarChar,  adParamInput,    64, vId)
      .Parameters.Append .CreateParameter("@Elog_Data",   adVarChar,  adParamInput,  8000, vData)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub



  '...get Transaction by CustId
  Function fElogById (vCustId)
    Dim oRs
    fElogById = ""
    sOpenDb
    vSql = " SELECT TOP 500 * FROM V5_Vubz.dbo.Elog " _
    		 & fIf(Len(vCustId) = 8, "WHERE Elog_CustId = '" & vCustId & "'", "") _ 
				 & " ORDER BY Elog_No DESC"
'   Response.Write vSql
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof        
      fElogById = fElogById & "<option value=" & oRs("Elog_No") & ">" & oRs("Elog_CustId") & " | " & fFormatSqlDate(oRs("Elog_Posted")) & " | " & oRs("Elog_Id") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb      
  End Function 


%>