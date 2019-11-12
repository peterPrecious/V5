<%
  Dim vTask_No, vTask_AcctId, vTask_Id, vTask_Order, vTask_Level, vTask_Child, vTask_Locked, vTask_Title, vTask_Desc, vTask_List
  Dim vTask_Crit1, vTask_Crit2, vTask_Crit3, vTask_Crit4, vTask_Crit5, vTask_DateStart, vTask_DateEnd, vTask_Active
  Dim vTask_Services
  Dim vTask_Eof

  '...Get Task Recordset
  Sub sGetTask_rs (vTaskAcctId, vTaskId)
    vSql = "SELECT * FROM Task WHERE Task_AcctId = '" & vTaskAcctId & "'"
    If vTaskId = 9990 Then
      vSql = vSql & " OR Task_AcctId = '0000' "
    ElseIf vTaskId <> 9999 Then
      vSql = vSql & " AND Task_Id = '" & vTaskId & "' "
    End If
'   sDebug
    sOpenDB    
    Set oRs = oDB.Execute(vSql)
  End Sub

  Sub sGetTask (vTaskAcctID, vTaskNo)
    vTask_Eof = False
    vSql = "SELECT * FROM Task WHERE Task_AcctId = '" & vTaskAcctId & "' AND Task_No = " & vTaskNo
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadTask
      vTask_Eof = True
    End If
    Set oRs = Nothing
    sCloseDb    
  End Sub

  Sub sReadTask
    vTask_AcctId          = oRs("Task_AcctId")
    vTask_No              = oRs("Task_No")
    vTask_Id              = oRs("Task_Id")
    vTask_Order           = oRs("Task_Order")
    vTask_Level           = oRs("Task_Level")
    vTask_Child           = oRs("Task_Child")
    vTask_Locked          = oRs("Task_Locked")
    vTask_Title           = oRs("Task_Title")
    vTask_Desc            = oRs("Task_Desc")
    vTask_Crit1           = oRs("Task_Crit1")
    vTask_Crit2           = oRs("Task_Crit2")
    vTask_DateStart       = oRs("Task_DateStart")
    vTask_DateEnd         = oRs("Task_DateEnd")
    vTask_List            = oRs("Task_List")
    vTask_Services        = oRs("Task_Services")
    vTask_Active          = oRs("Task_Active")
  End Sub

  Sub sExtractTask
    vTask_No              = Request.Form("Task_No")
    vTask_Id              = Request.Form("Task_Id")
    vTask_Order           = Request.Form("Task_Order")
    vTask_Level           = Request.Form("Task_Level")
    vTask_Child           = Request.Form("Task_Child")
    vTask_Locked          = Request.Form("Task_Locked")
    vTask_Title           = Request.Form("Task_Title")
    vTask_Desc            = Request.Form("Task_Desc")
    vTask_List            = Request.Form("Task_List")
    vTask_Active          = Request.Form("Task_Active")
    
    If fNoValue(vTask_Locked) Then vTask_Locked = 0
    If fNoValue(vTask_Child)  Then vTask_Child  = 0
    If fNoValue(vTask_Active) Then vTask_Active = 1
    
  End Sub
  
  Sub sInsertTask
    vSql = "INSERT INTO Task "
    vSql = vSql & "(Task_Id, Task_Order, Task_Level, Task_Level, Task_Child, Task_Title, Task_Desc, Task_Tasks)"
    vSql = vSql & " VALUES ('" & vTask_Id & "', " & vTask_Order & ", " & vTask_Level & ", " & vTask_Child & ", " & vTask_Locked & ", '" & fUnquote(vTask_Title) & "', '" & fUnquote(vTask_Desc) & "', '" & vTask_Tasks & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sUpdateTask
    vSql = "UPDATE Task SET"
    vSql = vSql & " Task_Title           = '" & fUnquote(vTask_Title)           & "', " 
    vSql = vSql & " Task_Desc            = '" & fUnquote(vTask_Desc)            & "', " 
    vSql = vSql & " Task_Level           =  " & vTask_Level                     & " , " 
    vSql = vSql & " Task_Child           =  " & vTask_Child                     & " , " 
    vSql = vSql & " Task_Locked          =  " & vTask_Locked                    & " , " 
    vSql = vSql & " Task_Tasks           = '" & fUnquote(vTask_Tasks)           & "'  " 
    vSql = vSql & " WHERE Task_No        = '" & vTask_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub
  
  Sub sDeleteTask
    vSql = "DELETE FROM Task WHERE Task_No = " & vTask_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


%>