<%

  Function sp5countryCodesDD (xxxCountry) '...dropdown for Ecom2Customer.asp
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp5countryCodes"
      .Parameters.Append .CreateParameter("@parm", adVarChar, adParamInput, 3, "*")
    End With
    Set oRs = oCmdApp.Execute()

    Dim i, char2, char3, country, selected
    i = vbCrLf 
    Do While Not oRs.Eof 
      country = oRs("country")
      char2 = oRs("char2")
      char3 = oRs("char3")
      selected = fIf(xxxCountry = char2, " selected", "")
      i = i & "          <option value=" & Chr(34) & char2 & Chr(34) & selected & ">" & country & "</option>" & vbCrLf
      oRs.MoveNext
    Loop 
    Set oRs = Nothing      
    Set oCmdApp = Nothing
    sCloseDbApp
    sp5countryCodesDD = i
  End Function



  Function sp5countryCodeElavon (xxxCountry) '...pass in a two char and get a 3 char back
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp5countryCodes"
      .Parameters.Append .CreateParameter("@parm", adVarChar, adParamInput, 3, xxxCountry)
    End With
    Set oRs = oCmdApp.Execute()
    sp5countryCodeElavon = oRs("char3")
    Set oRs = Nothing      
    Set oCmdApp = Nothing
    sCloseDbApp
  End Function


 %>