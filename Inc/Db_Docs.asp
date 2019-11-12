<%
  Dim vDocs_No, vDocs_AcctId, vDocs_TskHNo, vDocs_FileNo, vDocs_FileName, vDocs_Description, vDocs_MembNo, vDocs_Posted
  Dim vDocs_Eof

  
  '____ Docs  ________________________________________________________________________

  '...get an active document using the document no
  Sub sGetDocs (vKey)      
    vSql = "SELECT * FROM (Docs LEFT JOIN Memb ON Docs.Docs_MembNo = Memb.Memb_No) WHERE Docs.Docs_No = " & vKey
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then sReadDocs
    Set oRs = Nothing      
    sCloseDb
  End Sub

  '...get the current fields from the current record in the record set
  Sub sReadDocs
    vDocs_No           = oRs("Docs_No")
    vDocs_AcctId       = oRs("Docs_AcctId")
    vDocs_TskHNo       = oRs("Docs_TskHNo")
    vDocs_FileNo       = oRs("Docs_FileNo")
    vDocs_FileName     = oRs("Docs_FileName")
    vDocs_Description  = oRs("Docs_Description")
    vDocs_MembNo       = oRs("Docs_MembNo")
    vDocs_Posted       = oRs("Docs_Posted")
    sReadMemb
  End Sub

  '...read the Document list
  Sub sGetDocs_rs (vSort, vTskH_No)
    If fNoValue(vSort) Then vSort = "Docs_Posted DESC"
    '...if current TskH then get all items
    vSql = "SELECT * FROM (Docs LEFT JOIN Memb ON Docs.Docs_MembNo = Memb.Memb_No)"
    vSql = vSql & " WHERE Docs.Docs_AcctId = '" & svCustAcctId & "'"
    vSql = vSql & " AND Docs.Docs_TskHNo = " & vTskH_No
    vSql = vSql & " ORDER BY " & vSort
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
  End Sub
 
   '...get last upload
  Function fDocs_LastPosted
    fDocs_LastPosted = ""
    vSql = "SELECT TOP 1 Docs_Posted FROM Docs WHERE Docs_AcctId = '" & svCustAcctId & "' AND Docs_TskHNo = " & vTskH_No  & " AND Docs_MembNo <> " & svMembNo & " ORDER BY Docs_Posted DESC"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then Exit Function
    fDocs_LastPosted = oRs2("Docs_Posted")
    sCloseDb2    
  End Function
 
  Sub sDeleteDocs
    vSql = "DELETE FROM Docs WHERE Docs_No = " & vDocs_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub 
  
  '...get next Document no by inserting a new record and getting back the number
  Function fNextDocsNo

    vSql = "SET NOCOUNT ON "
    vSql = vSql & "INSERT INTO Docs "
    vSql = vSql & "(Docs_AcctId, Docs_TskHNo, Docs_FileName, Docs_Description, Docs_MembNo, Docs_Posted)"
    vSql = vSql & " VALUES ('" & svCustAcctId & "', " & vTskH_No & ", '" & vDocs_FileName & "', '" & vDocs_Description & "', " & svMembNo & ", '" & fFormatSqlDate (Now) & "')"
    vSql = vSql + " SELECT vFld=@@IDENTITY"
    vSql = vSql + " SET NOCOUNT OFF"

'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    vDocs_No = oRs("vFld")
    fNextDocsNo = Right("00000000" & vDocs_No, 8)
    Set oRs = Nothing      
    sCloseDb 
  End Function



%>