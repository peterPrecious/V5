<%
  '____ Path - used to log transactions that go through Patience.asp (mainly from Ecom2Checkout.asp and Ecom3Checout.asp  ___________________________________________

  Dim vPath_No, vPath_AcctId, vPath_MembNo, vPath_CatlNo, vPath_ProgNo, vPath_ModsNo, vPath_Posted
  Dim vPath_Eof


  Sub sReadPath
    vPath_No           = oRs("Path_No")
    vPath_AcctId       = oRs("Path_AcctId")
    vPath_MembNo       = oRs("Path_MembNo")
    vPath_CatlNo       = oRs("Path_CatlNo")
    vPath_ProgNo       = oRs("Path_ProgNo")
    vPath_ModsNo       = oRs("Path_ModsNo")
    vPath_Posted       = oRs("Path_Posted")
  End Sub


  '...Get a list of all Path for this account containing Path_CatlNo and Path_No
  Sub spPathById (vAcctId)
    sOpenCmd
    With oCmd
      .CommandText = "spPathSelect"
      .Parameters.Append .CreateParameter("@Path_AcctId", adChar,     adParamInput,     8, vAcctId)
      .Parameters.Append .CreateParameter("@Path_AcctId", adVarChar,  adParamInput,    64, vId)
    End With
    Set oRs = oCmd.Execute()
    If Not oRs.Eof Then 
      sReadPath
    Else
      sInitPath
    End If
    Set oRs = Nothing      
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...clear out any previous paths
  Sub spPathDelete (vMembNo)  
    sOpenCmd
    With oCmd
      .CommandText = "spPathDelete"
      .Parameters.Append .CreateParameter("@Path_MembNo",   adInteger,  adParamInput,      , vMembNo)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...insert current learning path
  Sub spPathInsert (vAcctId, vMembNo, vCatlNo, vProgNo, vModsNo)  
    sOpenCmd
    With oCmd
      .CommandText = "spPathInsert"
      .Parameters.Append .CreateParameter("@Path_AcctId",   adChar,     adParamInput,     4, vAcctId)
      .Parameters.Append .CreateParameter("@Path_MembNo",   adInteger,  adParamInput,      , vMembNo)
      .Parameters.Append .CreateParameter("@Path_CatlNo",   adInteger,  adParamInput,      , vCatlNo)
      .Parameters.Append .CreateParameter("@Path_ProgNo",   adInteger,  adParamInput,      , vProgNo)
      .Parameters.Append .CreateParameter("@Path_ModsNo",   adInteger,  adParamInput,      , vModsNo)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Sub




%>