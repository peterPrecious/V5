<%
  Dim vCaln_Date, vCaln_AcctId, vCaln_TskHNo, vCaln_Details
  Dim vCaln_Eof

  '____ Caln  ________________________________________________________________________

  '...Get Caln Record
  Sub sGetCaln (vCaln_Date, vTskHNo)
    vCaln_Eof = True
    vSql = "SELECT * FROM Caln WHERE Caln_Date= '" & fFormatSqlDate(vCaln_Date) & "' AND Caln_AcctId = '" & svCustAcctId & "' AND Caln_TskHNo = " & vTskHNo
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadCaln
      vCaln_Eof = False
    End If
    sCloseDb
  End Sub

   '...get last upload
  Function fCaln_LastPosted
    fCaln_LastPosted = ""
    vSql = "SELECT TOP 1 Caln_Posted FROM Caln WHERE Caln_AcctId = '" & svCustAcctId & "' AND Caln_TskHNo = " & vTskH_No  & " ORDER BY Caln_Posted DESC"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then Exit Function
    fCaln_LastPosted = oRs2("Caln_Posted")
    sCloseDb2    
  End Function

 '...get the current fields from the current record in the record set
  Sub sReadCaln
    vCaln_Date            = fFormatSqlDate(oRs("Caln_Date"))
    vCaln_TskHNo          = oRs("Caln_TskHNo")
    vCaln_Details         = oRs("Caln_Details")
  End Sub
  
  '...insert Caln
  Sub sInsertCaln
    vSql = "INSERT INTO Caln "
    vSql = vSql & "(Caln_Date, Caln_AcctId, Caln_TskHNo, Caln_Details)"
    vSql = vSql & " VALUES ('" & fFormatSqlDate(vCaln_Date) & "', '" & svCustAcctId & "', " & vTskH_No & ", '" & vCaln_Details & "')"
'   sDebug
    sOpenDb
    On Error Resume Next
    oDb.Execute(vSql)
    sCloseDb
    If Error <> 0 Then sUpdateCaln
  End Sub  
  
  Sub sUpdateCaln
    vSql = "UPDATE Caln SET"
    vSql = vSql & " Caln_Details =  '" & vCaln_Details & "' " 
    vSql = vSql & " WHERE Caln_Date= '" & fFormatSqlDate(vCaln_Date) & "' AND Caln_AcctId = '" & svCustAcctId & "' AND Caln_TskHNo = " & vTskH_No
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub  
 
%>