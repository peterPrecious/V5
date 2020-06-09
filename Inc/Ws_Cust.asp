<%
  '____ Cust  ________________________________________________________________________

  Dim vCust_Id, vCust_AcctId, vCust_Title, vCust_Lang, vCust_Active, vCust_Auth, vCust_Prgms
  Dim vCust_Eof

  '...Get Cust Recordset
  Sub sGetCust (vCustId)


    vSql = "SELECT Cust_Id, Cust_AcctID, Cust_Title, Cust_Lang, Cust_Programs, Cust_Active, Cust_Auth FROM Cust WHERE Cust_Id= '" & vCustId & "'"
    sOpenDB
    Set oRs = oDB.Execute(vSql)
    If Not oRs.Eof Then 
      sReadCust
      vCust_Eof = False
      If vCust_Active = 0 then vCust_Eof = True
      If vCust_Auth   = 0 then vCust_Eof = True
      vResponseNamePair = "vCust_Id=" & vCust_Id _
       & "&vCust_AcctId=" & vCust_AcctId _
       & "&vCust_Title="  & vCust_Title _
       & "&vCust_Lang="   & vCust_Lang _
       & "~||~vCust_Prgms="  & vCust_Prgms   '... ~||~ keep vCust_Prgms seperate to parse out later
    Else
      vCust_Eof = True
    End If
    Set oRs = Nothing
    sCloseDB    
  End Sub

  Sub sReadCust
    vCust_Id                = oRs("Cust_Id")
    vCust_AcctId            = oRs("Cust_AcctId")
    vCust_Title             = oRs("Cust_Title")
    vCust_Lang              = oRs("Cust_Lang")
    vCust_Prgms             = oRs("Cust_Programs")
    vCust_Active            = oRs("Cust_Active")
    vCust_Auth              = oRs("Cust_Auth")
  End Sub

%>