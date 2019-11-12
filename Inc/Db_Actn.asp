<%
  Dim vActn_No, vActn_AcctId, vActn_TskHNo, vActn_Item, vActn_Completed, vActn_Due, vActn_Posted, vActn_Owner
  Dim vActn_Eof

  '____ Actn  ________________________________________________________________________

  '...get an action item using the item no  
  Sub sGetActn (vNo)
    vSql = "SELECT * FROM (Actn LEFT JOIN Memb ON Actn.Actn_Owner = Memb.Memb_No) WHERE Actn.Actn_No = " & vNo
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then sReadActn
    Set oRs = Nothing      
    sCloseDb
  End Sub

  '...get a message set
  Sub sGetActn_rs
    vSql = "SELECT * FROM (Actn LEFT JOIN Memb ON Actn.Actn_Owner = Memb.Memb_No) "
    vSql = vSql & " WHERE Actn.Actn_AcctId = '" & svCustAcctId & "' AND Actn.Actn_TskHNo = " & vTskH_No
    vSql = vSql & " ORDER BY " & vSort
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
  End Sub

   '...get last upload
  Function fActn_LastPosted
    fActn_LastPosted = ""
    vSql = "SELECT TOP 1 Actn_Posted FROM Actn WHERE Actn_AcctId = '" & svCustAcctId & "' AND Actn_TskHNo = " & vTskH_No  & " AND Actn_Owner = " & svMembNo & " ORDER BY Actn_Posted DESC"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then Exit Function
    fActn_LastPosted = oRs2("Actn_Posted")
    sCloseDb2    
  End Function
  
  Sub sReadActn
    vActn_No          = oRs("Actn_No")
    vActn_AcctId      = oRs("Actn_AcctId")
    vActn_TskHNo      = oRs("Actn_TskHNo")
    vActn_Owner       = oRs("Actn_Owner")
    vActn_Posted      = oRs("Actn_Posted")
    vActn_Completed   = oRs("Actn_Completed")
    vActn_Due         = oRs("Actn_Due")
    vActn_Item        = oRs("Actn_Item")
    sReadMemb
  End Sub
  
  '...Add an Action Item
  Sub sInsertActn
    If fNoValue(vActn_Completed) Then vActn_Completed = 0
    vSql = "INSERT INTO Actn "
    vSql = vSql & "(Actn_AcctId, Actn_TskHNo, Actn_Item, Actn_Due, Actn_Completed, Actn_Owner)"
    vSql = vSql & " VALUES ('" & svCustAcctId & "', " & vTskH_No & ", '" & fUnquote(vActn_Item) & "','" & vActn_Due & "', " & vActn_Completed & ", " & vActn_Owner & ")"
'   sDebug    
    sOpenDb
    oDb.Execute(vSql)
    Set oRs = Nothing      
    sCloseDb
  End Sub

  Sub sUpdateActn
    '...Must leave something
    If Len(Trim(vActn_Item)) = 0 Then vActn_Item = "...empty"
    If fNoValue(vActn_Completed) Then vActn_Completed = 0
    vSql = "UPDATE Actn SET"
    vSql = vSql & " Actn_Due         = '" & vActn_Due            & "', " 
    vSql = vSql & " Actn_Item        = '" & fUnquote(vActn_Item) & "', " 
    vSql = vSql & " Actn_Completed   =  " & vActn_Completed      & " , "
    vSql = vSql & " Actn_Owner       =  " & vActn_Owner
    vSql = vSql & " WHERE Actn_No    =  " & vActn_No
 '  sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sExtractActn
    vActn_Posted       = Request.Form("vActn_Posted")
    vActn_Item         = Request.Form("vActn_Item")
    vActn_Due          = Request.Form("vActn_Due")
    vActn_Completed    = Request.Form("vActn_Completed")
    vActn_Owner        = Request.Form("vActn_Owner")
  End Sub

  '...Delete Action Item
  Sub sDeleteActn
    sOpenDb
    vSql = "DELETE FROM Actn WHERE Actn_No = " & vActn_No
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...Delete Action Item
  Sub sCompletedActn
    sOpenDb
    vSql = "UPDATE Actn SET Actn_Completed = 1 WHERE Actn_No = " & vActn_No
'   sDebug
    sOpenDB
    oDB.Execute(vSql)
    sCloseDB
  End Sub

%>