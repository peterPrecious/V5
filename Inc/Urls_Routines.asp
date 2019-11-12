<%
  '____ Urls ________________________________________________________________________

  Function fGetUrls (vUrls_No)
    vSql = "SELECT * FROM Urls WHERE Urls_No = " & vUrls_No
    On Error Resume Next
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)    
    fGetUrls = oRs2("Urls_Address")
    Set oRs2 = Nothing      
    sCloseDb2
  End Function

  '...insert url and return url no
  Function fCreateUrls (vUrl)
    vSQL = "INSERT INTO Urls (Urls_Address) VALUES ('" & vUrl & "')"
    sOpenDb2
    oDb2.Execute(vSql)
    vSql = "SELECT TOP 1 Urls_No FROM Urls ORDER BY Urls_No DESC"
    Set oRs2 = oDb2.Execute(vSql)
    fCreateUrls = oRs2("Urls_No")
    Set oRs2 = Nothing      
    sCloseDb2    
  End Function

%>