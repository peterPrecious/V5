<%
  Dim vTskH_No, vTskH_AcctId, vTskH_Id, vTskH_OrderL0, vTskH_Order, vTskH_Level, vTskH_Child, vTskH_Locked, vTskH_Title, vTskH_Desc, vTskH_Active, vTskH_Password, vTskH_Collapse, vTskH_Lang
  Dim vTskH_AccessLevel, vTskH_AccessIds, vTskH_Criteria, vTskH_Group2, vTskH_CustIds, vTskH_DateStart, vTskH_DateEnd
  Dim vTskH_Notes, vTskH_Dialogue, vTskH_ActionItems, vTskH_Repository, vTskH_EmailAlert, vTskH_Calendar

  Dim vTskH_Eof

  '...Get TskH Recordset
  Sub sGetTskH_rs (vTskH_AcctId, vTskH_Id)
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' "
    '...extract all tasks and templates
    If vTskH_Id = 9990 Then
      vSql = vSql & " OR TskH_AcctId = '0000' "
   '...just extract tasks for this accout
    ElseIf vTskH_Id <> 9999 Then
      vSql = vSql & " AND TskH_Id = '" & vTskH_Id & "' "
    End If
'   sDebug
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
  End Sub

  Sub sGetTskH (vTskH_AcctId, vTskH_No)
    vTskH_Eof = False
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_No = " & vTskH_No
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadTskH
      vTskH_Eof = True
    End If
    Set oRs = Nothing
    sCloseDb    
  End Sub

  Sub sGetTskH0 (vTskH_AcctId, vTskH_Id)
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id & " AND TskH_Level = 0"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then sReadTskH
    Set oRs = Nothing
    sCloseDb    
  End Sub

  Sub sReadTskH
    vTskH_AcctId          = oRs("TskH_AcctId")
    vTskH_No              = oRs("TskH_No")
    vTskH_Id              = oRs("TskH_Id")
    vTskH_Lang            = oRs("TskH_Lang")
    vTskH_OrderL0         = oRs("TskH_OrderL0")
    vTskH_Order           = oRs("TskH_Order")
    vTskH_Level           = oRs("TskH_Level")
    vTskH_Child           = oRs("TskH_Child")
    vTskH_Locked          = oRs("TskH_Locked")
    vTskH_Title           = oRs("TskH_Title")
    vTskH_Desc            = oRs("TskH_Desc")
    vTskH_AccessLevel     = oRs("TskH_AccessLevel")
    vTskH_AccessIds       = oRs("TskH_AccessIds")
    vTskH_Criteria        = oRs("TskH_Criteria")
    vTskH_Group2          = oRs("TskH_Group2")
    vTskH_CustIds         = oRs("TskH_CustIds")
    vTskH_DateStart       = oRs("TskH_DateStart")
    vTskH_DateEnd         = oRs("TskH_DateEnd")
    vTskH_Notes           = oRs("TskH_Notes")
    vTskH_Dialogue        = oRs("TskH_Dialogue")
    vTskH_ActionItems     = oRs("TskH_ActionItems")
    vTskH_Repository      = oRs("TskH_Repository")
    vTskH_EmailAlert      = oRs("TskH_EmailAlert")
    vTskH_Calendar        = oRs("TskH_Calendar")
    vTskH_Active          = oRs("TskH_Active")
    vTskH_Password        = oRs("TskH_Password")
    vTskH_Collapse        = oRs("TskH_Collapse")
  End Sub

  Sub sExtractTskH
    vTskH_No              = Request.Form("vTskH_No")
    vTskH_Id              = Request.Form("vTskH_Id")
    vTskH_Lang            = fDefault(Request.Form("vTskH_Lang"),"XX")
    vTskH_Order           = Request.Form("vTskH_Order")
    vTskH_Level           = Request.Form("vTskH_Level")
    vTskH_Child           = Request.Form("vTskH_Child")
    vTskH_Locked          = Request.Form("vTskH_Locked")
    vTskH_Title           = fUnquote(Request.Form("vTskH_Title"))
    vTskH_Desc            = fUnquote(Request.Form("vTskH_Desc"))
    vTskH_AccessLevel     = Request.Form("vTskH_AccessLevel")
    vTskH_AccessIds       = Trim(Ucase(Request.Form("vTskH_AccessIds")))
    vTskH_Criteria        = Replace(Request.Form("vTskH_Criteria"), ",", "")
    vTskH_Group2          = Request.Form("vTskH_Group2")
    vTskH_CustIds         = Request.Form("vTskH_CustIds")
    vTskH_DateStart       = Request.Form("vTskH_DateStart")
    vTskH_DateEnd         = Request.Form("vTskH_DateEnd")
    vTskH_Notes           = Request.Form("vTskH_Notes")
    vTskH_Dialogue        = Request.Form("vTskH_Dialogue")
    vTskH_ActionItems     = Request.Form("vTskH_ActionItems")
    vTskH_Repository      = Request.Form("vTskH_Repository")
    vTskH_EmailAlert      = Request.Form("vTskH_EmailAlert")
    vTskH_Calendar        = Request.Form("vTskH_Calendar")
    vTskH_Active          = Request.Form("vTskH_Active")
    vTskH_Collapse        = Request.Form("vTskH_Collapse")
    vTskH_Password        = fEncode(fNoquote(Ucase(Trim(Request.Form("vTskH_Password")))))

    If fNoValue(vTskH_Criteria)    Then vTskH_Criteria    = 0    
    If fNoValue(vTskH_Locked)      Then vTskH_Locked      = 0
    If fNoValue(vTskH_Child)       Then vTskH_Child       = 0
    If fNoValue(vTskH_Active)      Then vTskH_Active      = 0
    If fNoValue(vTskH_AccessLevel) Then vTskH_AccessLevel = 2
    If fNoValue(vTskH_Dialogue)    Then vTskH_Dialogue    = 0
    If fNoValue(vTskH_ActionItems) Then vTskH_ActionItems = 0
    If fNoValue(vTskH_Notes)       Then vTskH_Notes       = 0
    If fNoValue(vTskH_Repository)  Then vTskH_Repository  = 0
    If fNoValue(vTskH_EmailAlert)  Then vTskH_EmailAlert  = 0
    If fNoValue(vTskH_Calendar)    Then vTskH_Calendar    = 0
    If fNoValue(vTskH_Collapse)    Then vTskH_Collapse    = 0

    If Not IsDate(vTskH_DateStart) Then vTskH_DateStart   = "" 
    If Not IsDate(vTskH_DateEnd  ) Then vTskH_DateEnd     = "" 

  End Sub
  
  Sub sInsertTskH
    vSql = "INSERT INTO TskH "
    vSql = vSql & "(TskH_AcctId, TskH_Id, TskH_Lang, TskH_OrderL0, TskH_Order, TskH_Level, TskH_Child, TskH_Locked, TskH_Title, TskH_Desc, TskH_AccessLevel, TskH_AccessIds, TskH_Criteria, TskH_CustIds, TskH_Notes, TskH_Dialogue, TskH_ActionItems, TskH_Repository, TskH_EmailAlert, TskH_Calendar, TskH_Collapse, TskH_Password)"
    vSql = vSql & " VALUES ('" & vTskH_AcctId & "', " & vTskH_Id & ", '" & vTskH_Lang & "', " & vTskH_OrderL0 & ", " & vTskH_Order & ", " & vTskH_Level & ", " & fSqlBoolean(vTskH_Child) & ", " & fSqlBoolean(vTskH_Locked) & ", '" & vTskH_Title & "', '" & vTskH_Desc & "', " & vTskH_AccessLevel & ", '" & vTskH_AccessIds & "', '" & vTskH_Criteria & "', '" & vTskH_CustIds & "', " & fSqlBoolean(vTskH_Notes) & ",  " & fSqlBoolean(vTskH_Dialogue) & ",  " & fSqlBoolean(vTskH_ActionItems) & ",  " & fSqlBoolean(vTskH_Repository) & ",  " & fSqlBoolean(vTskH_EmailAlert) & ",  " & fSqlBoolean(vTskH_Calendar) & ",  " & fSqlBoolean(vTskH_Collapse) & ", '" & vTskH_Password & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sUpdateTskH
    vSql = "UPDATE TskH SET"
    vSql = vSql & " TskH_Lang            = '" & vTskH_Lang                      & "', " 
    vSql = vSql & " TskH_Title           = '" & vTskH_Title                     & "', " 
    vSql = vSql & " TskH_Desc            = '" & vTskH_Desc                      & "', " 
    vSql = vSql & " TskH_Level           =  " & vTskH_Level                     & " , " 
    vSql = vSql & " TskH_Child           =  " & vTskH_Child                     & " , " 
    vSql = vSql & " TskH_Locked          =  " & vTskH_Locked                    & " , " 

    vSql = vSql & " TskH_AccessLevel     =  " & vTskH_AccessLevel               & " , " 
    vSql = vSql & " TskH_AccessIds       = '" & vTskH_AccessIds                 & "', " 
    vSql = vSql & " TskH_Criteria        = '" & vTskH_Criteria                  & "', " 
    vSql = vSql & " TskH_Group2          =  " & vTskH_Group2                    & " , " 
    vSql = vSql & " TskH_CustIds         = '" & vTskH_CustIds                   & "', " 
    vSql = vSql & " TskH_DateStart       = '" & vTskH_DateStart                 & "', " 
    vSql = vSql & " TskH_DateEnd         = '" & vTskH_DateEnd                   & "', " 
    vSql = vSql & " TskH_Active          =  " & vTskH_Active                    & " , " 
    vSql = vSql & " TskH_Password        = '" & vTskH_Password                  & "', " 
'   vSql = vSql & " TskH_Collapse        =  " & vTskH_Collapse                  & " , " 

    vSql = vSql & " TskH_Notes           =  " & vTskH_Notes                     & " , " 
    vSql = vSql & " TskH_Dialogue        =  " & vTskH_Dialogue                  & " , " 
    vSql = vSql & " TskH_ActionItems     =  " & vTskH_ActionItems               & " , " 
    vSql = vSql & " TskH_Repository      =  " & vTskH_Repository                & " , " 
    vSql = vSql & " TskH_Calendar        =  " & vTskH_Calendar                  & " , " 
    vSql = vSql & " TskH_EmailAlert      =  " & vTskH_EmailAlert

    vSql = vSql & " WHERE TskH_No        =  " & vTskH_No
    sOpenDb
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sDeleteTskH_rs (vTskH_AcctId, vTskH_Id)

    '...first delete any assets
    vSql = "SELECT TskH_No FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      vSql = "DELETE FROM TskD WHERE TskD_No = " & oRs("TskH_No")
      oDb.Execute(vSql)
      oRs.MoveNext
    Loop
    Set oRs = Nothing
  
    '...now delete the task
    vSql = "DELETE FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    oDb.Execute(vSql)
    sCloseDb
  End Sub
  
  Sub sDeleteTskH
    vSql = "DELETE FROM TskH WHERE TskH_No = " & vTskH_No
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sActivateTskH (vTskH_AcctId, vTskH_Id)
    vSql = "UPDATE TskH SET TskH_Active = 1 WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sInActivateTskH (vTskH_AcctId, vTskH_Id)
    vSql = "UPDATE TskH SET TskH_Active = 0 WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...this sets the expand/collapse field for all items in this task
  Sub sCollapseTskH (vTskH_AcctId, vTskH_Id, vTskH_Collapse)
    vSql = "UPDATE TskH SET TskH_Collapse = " & vTskH_Collapse & " WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sLockTskH (vTskH_No)
    vSql = "UPDATE TskH SET TskH_Locked = 1 WHERE TskH_No = " & vTskH_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sUnlockTskH (vTskH_No)
    vSql = "UPDATE TskH SET TskH_Locked = 0 WHERE TskH_No = " & vTskH_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sResetLockTskH (vTskH_No)
    vSql = "UPDATE TskH SET TskH_Locked = 0 WHERE TskH_No = " & vTskH_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sCloneTskH (vTskH_AcctId, vTskH_Id)
    '...this clones an entire task set
    Dim vNextId, vNextNo
    vNextId = fNextTskH_Id
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      sReadTskH
      '...unquote because we're bypassing extract
      vTskH_Title  = fUnquote(vTskH_Title)
      vTskH_Desc   = fUnquote(vTskH_Desc)     
      vNextNo      = fInsertNextTskH (svCustAcctId, vNextId)
      '...copy assets (active only when the corresponding delete works)
      sInsertNextTskD vTskH_No, vNextNo
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  End Sub


  '...same as clone except the AcctId become "0000"
  Sub sTemplateTskH (vTskH_AcctId, vTskH_Id)
    '...this clones an entire task set
    Dim vNextId, vNextNo
    vNextId = fNextTskH_Id
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      sReadTskH
      '...unquote because we're bypassing extract
      vTskH_Title  = fUnquote(vTskH_Title)
      vTskH_Desc   = fUnquote(vTskH_Desc)     
      vNextNo      = fInsertNextTskH ("0000", vNextId)
      '...copy assets (active only when the corresponding delete works)
      sInsertNextTskD vTskH_No, vNextNo
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  End Sub


  '...same as clone except clones all items from Old Acct to New Acct
  Sub sCloneSite (vOldAcctId, vNewAcctId)
    '...this clones an entire task set
    Dim vNextId, vNextNo
    vNextId = fNextTskH_Id
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vOldAcctId & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      sReadTskH
      '...unquote because we're bypassing extract
      vTskH_Title  = fUnquote(vTskH_Title)
      vTskH_Desc   = fUnquote(vTskH_Desc)     
      vNextNo      = fInsertNextTskH (vNewAcctId, vNextId)
      '...copy assets (active only when the corresponding delete works)
      sInsertNextTskD vTskH_No, vNextNo
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  End Sub




  Sub sCloneTskHbyNo (vTskH_Id, vTskH_No)
    '...this clones an individual task item
    vSql = "SELECT * FROM TskH WHERE TskH_No = " & vTskH_No
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    sReadTskH
    '...unquote because we're bypassing extract
    vTskH_Title           = fUnquote(vTskH_Title)
    vTskH_Desc            = fUnquote(vTskH_Desc)     
    sCloseDb    
'   vTskH_Order = fNextOrderNo (vTskH_Id)
    vTskH_Order = vTskH_Order + .5
    sInsertTskH     
    Set oRs = Nothing

    '...reorder the task list from cloned value to the end to make order an integer
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & svCustAcctId & "' AND TskH_Id = '" & vTskH_Id & "' AND TskH_Order >= " & vTskH_Order & " ORDER BY TskH_Order DESC "
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      sReadTskH
      vSql =  "UPDATE TskH SET TskH_Order = CAST(" & vTskH_Order + 1 & " AS INT) WHERE TskH_No =  " & vTskH_No
'     sDebug
      oDb.Execute(vSql)
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
  End Sub


  Sub sAddTskH (vTskH_Id, vTskH_No, vNextOrder, vLevel)
    vSql = "SELECT * FROM TskH WHERE TskH_No = " & vTskH_No
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    sReadTskH
    sCloseDb    
    '...override values
    vTskH_Order = Cint(vNextOrder)
    vTskH_Level = Cint(vLevel)
    vTskH_Title = "New item..."
    vTskH_Desc  = fUnquote(vTskH_Desc)     
    sInsertTskH     
    Set oRs = Nothing
  End Sub
  
  '...this gives the next available order no
  Function fNextOrderNo (vTskH_Id)
    vSql = "SELECT MAX(TskH_Order) AS NextOrderNo FROM TskH WHERE TskH_Id = " & vTskH_Id
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    fNextOrderNo = oRs2("NextOrderNo") + 1
    sCloseDb2
  End Function

  '...this gives the next available id
  Function fNextTskH_Id
    vSql = "SELECT MAX(TskH_Id) AS NextId FROM TskH"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    fNextTskH_Id = oRs2("NextId") + 1
    sCloseDb2
  End Function

  '...this gives the next id in the account after given Id to move down a task group
  Function fNextInAcctTskH_Id (vTskH_AcctId, vTskH_Id)
    vSql = "SELECT MIN(TskH_Id) AS NextId FROM TskH WHERE (TskH_Id > " & vTskH_Id & ") AND (TskH_AcctId = '" & vTskH_AcctId & "')"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    fNextInAcctTskH_Id = oRs2("NextId")
    If fNoValue(fNextInAcctTskH_Id) Then
      fNextInAcctTskH_Id = 0
    End If
    sCloseDb2
  End Function

  Function fInsertNextTskH (vAcctId, vId)
    vSql = "SET NOCOUNT ON "
    vSql = vSql & "INSERT INTO TskH "
    vSql = vSql & "(TskH_AcctId, TskH_Id, TskH_Active, TskH_Lang, TskH_Order, TskH_Level, TskH_Child, TskH_Locked, TskH_Password, TskH_Title, TskH_Desc, TskH_AccessLevel, TskH_Criteria, TskH_Group2, TskH_CustIds, TskH_Notes, TskH_Dialogue, TskH_ActionItems, TskH_Repository, TskH_EmailAlert, TskH_Calendar)"
    vSql = vSql & " VALUES ('" & vAcctId & "', " & vId & ", " & fSqlBoolean (vTskH_Active) & ", '" & vTskH_Lang & "', " & vTskH_Order & ", " & vTskH_Level & ", " & fSqlBoolean(vTskH_Child) & ", " & fSqlBoolean(vTskH_Locked) & ", '" & vTskH_Password & "', '" & vTskH_Title & "', '" & vTskH_Desc & "', " & vTskH_AccessLevel & ", '" & vTskH_Criteria & "', " & vTskH_Group2 & ", '" & vTskH_CustIds & "', " & fSqlBoolean(vTskH_Notes) & ",  " & fSqlBoolean(vTskH_Dialogue) & ",  " & fSqlBoolean(vTskH_ActionItems) & ",  " & fSqlBoolean(vTskH_Repository) & ",  " & fSqlBoolean(vTskH_EmailAlert) & ",  " & fSqlBoolean(vTskH_Calendar) & ")"
    vSql = vSql + " SELECT vFld=@@IDENTITY"
    vSql = vSql + " SET NOCOUNT OFF"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    fInsertNextTskH = oRs2("vFld")
    sCloseDb2
  End Function

  Sub sTaskOrderUp
    sOpenDb
    vOrderNo = Cint(sInput("vTskH_Order"))
    '...give desired row a dummy sort order
    vSql = "UPDATE TskH SET TskH_Order = 999999 WHERE TskH_No = " & sInput("vTskH_No")
    oDb.Execute(vSql)
    '...get previous row and set it to current row
    vSql = "UPDATE TskH SET TskH_Order = " & vOrderNo & " WHERE TskH_AcctId = '" & sInput("vTskH_AcctId") & "' AND TskH_Id = " & sInput("vTskH_Id") & " AND TskH_Order = " & vOrderno - 1 
    oDb.Execute(vSql)
    '...get current row and set it to previous row
    vSql = "UPDATE TskH SET TskH_Order = " & vOrderNo - 1 & " WHERE TskH_No = " & sInput("vTskH_No")
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sTaskOrderDown
    sOpenDb
    vOrderNo = Cint(sInput("vTskH_Order"))
    '...give desired row a dummy sort order
    vSql = "UPDATE TskH SET TskH_Order = 999999 WHERE TskH_No = " & sInput("vTskH_No")
    oDb.Execute(vSql)
    '...get next row and set it to current row
    vSql = "UPDATE TskH SET TskH_Order = " & vOrderNo  & " WHERE TskH_AcctId = '" & sInput("vTskH_AcctId") & "' AND TskH_Id = " & sInput("vTskH_Id") & " AND TskH_Order = " & vOrderno + 1 
    oDb.Execute(vSql)
    '...get current row and set it to next row
    vSql = "UPDATE TskH SET TskH_Order = " & vOrderNo + 1 & " WHERE TskH_No = " & sInput("vTskH_No")
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Function fMaxTskH_Order
    vSql = "SELECT MAX(TskH_Order) AS MaxOrder FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    fMaxTskH_Order = oRs("MaxOrder") 
    sCloseDb
  End Function

  Sub sFlagTaskChildren (vTskH_AcctId, vTskH_Id)
    Dim vTskH_Level_Prev
    vTskH_Level_Prev = 0
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id & " ORDER BY TskH_Order DESC "
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      sReadTskH
'     sDebug vTskH_Level, vTskH_Child
      If vTskH_Level < vTskH_Level_Prev and vTskH_Level > 0 Then
        If vTskH_Child <> True Then
          oDb.Execute("UPDATE TskH SET TskH_Child = 1 WHERE TskH_No =  " & vTskH_No)
        End If 
      Else
        If vTskH_Child <> False Then
          oDb.Execute("UPDATE TskH SET TskH_Child = 0 WHERE TskH_No =  " & vTskH_No)
        End If 
      End If
      vTskH_Level_Prev = vTskH_Level
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  End Sub  
  
  Function fGetParentTitle (vTskH_AcctId, vTskH_Id, vTskH_Order)
    Dim vTskH_Level_Prev
    vTskH_Level_Prev = 0
    fGetParentTitle = ""
    vSql = "SELECT TskH_Title FROM TskH WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & vTskH_Id & " AND TskH_Order < " & vTskH_Order & " AND TskH_Level = 1 ORDER BY TskH_Order DESC "
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then
      fGetParentTitle = oRs("TskH_Title")
    End If
    Set oRs = Nothing
    sCloseDb    
  End Function
  
  Function fNoTasks

    Dim vTskH_IdOk '...keep a copy of the latest task that is ok, because if there's only one, we need to pass through the valid Id, not the last one checked
  
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & svCustAcctId & "' AND TskH_Level = 0 AND TskH_Active = 1"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    fNoTasks = 0
    Do While Not oRs.Eof 
      sReadTskH
      If fTaskFilterOk Then
        fNoTasks = fNoTasks + 1
        vTskH_IdOk = vTskH_Id
      End If  
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    

    '...if only 1 valid task then
    If fNoTasks = 1 Then
      vTskH_Id = vTskH_IdOk
    End If

  End Function




  Function fTaskFilterOk    
    Dim aTskHCriteria, aMembCriteria
    fTaskFilterOk = False

    '...learner only ok
    If vMemb_Level <> 2 And vMemb_Level <> 5 And vTskH_AccessLevel = 1 Then
      Exit Function
    End If

    '...if the date ok
    If Not fNoValue(vTskH_DateStart) Then
      If Year(vTskH_DateStart) > 2000 Then
        If Now < vTskH_DateStart Then
          Exit Function
        End If
      End If
    End If
    If Not fNoValue(vTskH_DateEnd) Then
      If Year(vTskH_DateEnd) > 2000 Then
        If Now > vTskH_DateEnd Then
          Exit Function
        End If
      End if
    End If

    '...if customer id's ok
    If Len(Trim(vTskH_CustIds)) > 0 Then
      If Instr(Ucase(vTskH_CustIds), svCustId) = 0 Then Exit Function
    End If

    '...if learner id's ok and then check if level is ok
    If Len(Trim(vTskH_AccessIds)) > 0 Then
      If Instr(Ucase(vTskH_AccessIDs), svMembId) = 0 And svMembLevel < 5 Then Exit Function
    Else
      If vMemb_Level < vTskH_AccessLevel Then Exit Function
    End If

    '...is language ok
    If vTskH_Lang <> "XX" Then
      If vTskH_Lang <> svLang Then Exit Function
    End If

    '...any group2 values? (run this before Group1 as next group is tricky)
    If vTskH_Group2 <> 0 And vMemb_Group2 <> 0 Then
      If vTskH_Group2 <> vMemb_Group2 Then Exit Function
    End If

    '...any group 1 values?  ok to pass thru if vTskH_Criteria = 0 or the member criteria (group 2 value) = 0
    fTaskFilterOk = True
    If (Not fNoValue(vTskH_Criteria) And Trim(vTskH_Criteria) <> "0") And vMemb_Criteria <> "0" Then 
      aTskHCriteria = Split(vTskH_Criteria, " ")
      aMembCriteria = Split(vMemb_Criteria, " ")
      For i = 0 to Ubound(aTskHCriteria)
        For j = 0 to Ubound(aMembCriteria)
          If aMembCriteria(j) = aTskHCriteria(i) Then
            Exit Function
          End If
        Next        
      Next      
      fTaskFilterOk = False
    End If

  End Function


  Function fTaskOptions
    vSql = "SELECT * FROM TskH WHERE TskH_AcctId = '" & svCustAcctId & "' AND TskH_Level = 0"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    fTaskOptions = ""
    Do While Not oRs.Eof 
      sReadTskH
      If fTaskFilterOk Then
        fTaskOptions = fTaskOptions & "          <option value='" & vTskH_Id & "'>" & vTskH_Title & "</option>" & vbCrLf
      End If
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  End Function


  Function fOrderNo (vTskH_AcctId, vTskH_Id, vTskH_No)
    vSql = "SELECT * FROM TskH WHERE TskH_No = " & vTskH_No
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    fNextInAcctTskH_Id = oRs2("NextId")
    If fNoValue(fNextInAcctTskH_Id) Then
      fNextInAcctTskH_Id = 0
    End If
    sCloseDb2
  End Function
 
 
  Sub sShiftOrder (vTskH_AcctId, vTskH_Id, vTskHNo, vAction)

    Dim aRs(), vCurrNo, vSeekNo
    i = 0
    vCurrNo = 0
    vSeekNo = vTskHNo
    
    '...put the recordset into an array and flag the current vTskHNo so we can find the value before or after
    sGetTskH_rs vTskH_AcctId, vTskH_Id
    Do While Not oRs.Eof 
      sReadTskH
      i = i + 1
      ReDim Preserve aRs(2, i)
      aRs(1, i) = vTskH_No      
      aRs(2, i) = vTskH_Order
      If vTskH_No = vSeekNo Then
        vCurrNo = i
      End If
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
    
    '...if "up" and top record or "down" and bottom record then return without shifting
    If (vAction = "up" and vCurrNo = 1) Or (vAction = "down" and vCurrNo = ubound(aRs,2)) Then
      Exit Sub
    End If
    
    '...now that we have the two nos, swap their order value
    sOpenDb
    '...store the current order
    vSql = "UPDATE TskH Set TskH_Order = " & -999999         & " WHERE TskH_No = " & aRs(1, vCurrNo)
    Set oRs = oDb.Execute(vSql)
    If vAction = "down" Then
      vSql = "UPDATE TskH Set TskH_Order = " & aRs(2, vCurrNo) & " WHERE TskH_No = " & aRs(1, vCurrNo + 1)
      Set oRs = oDb.Execute(vSql)
      '...set the next order to the current order
      vSql = "UPDATE TskH Set TskH_Order = " & aRs(2, vCurrNo + 1) & " WHERE TskH_No = " & aRs(1, vCurrNo)
      Set oRs = oDb.Execute(vSql)
    Else
      vSql = "UPDATE TskH Set TskH_Order = " & aRs(2, vCurrNo) & " WHERE TskH_No = " & aRs(1, vCurrNo - 1)
      Set oRs = oDb.Execute(vSql)
      '...set the next order to the current order
      vSql = "UPDATE TskH Set TskH_Order = " & aRs(2, vCurrNo - 1) & " WHERE TskH_No = " & aRs(1, vCurrNo)
      Set oRs = oDb.Execute(vSql)
    End If
    sCloseDb
  End Sub
  
  Sub sShiftTask (vTskH_AcctId, Id1, Id2)
    sOpenDb2
    vSql = "UPDATE TskH Set TskH_Id = " & -99 & " WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & Id1
    Set oRs2 = oDb2.Execute(vSql)
    vSql = "UPDATE TskH Set TskH_Id = " & Id1 & " WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = " & Id2
    Set oRs2 = oDb2.Execute(vSql)
    vSql = "UPDATE TskH Set TskH_Id = " & Id2 & " WHERE TskH_AcctId = '" & vTskH_AcctId & "' AND TskH_Id = -99"
    Set oRs2 = oDb2.Execute(vSql)
    sCloseDb2
  End Sub   
  
  Sub sTskH_ClearSession  
    For Each vFld In Session.Contents
      If Len(vFld) > 5 Then
        If Right(vFld, 4) = "Tree" Then
          Session(vFld) = ""
        End If
      End If
    Next
  End Sub
%>