  <%
    '...these are similiar to other functions available in ModuleStatusRoutines, etc.

    Function fMembName (vMembNo)
      fMembName = ""
      vSql = "SELECT Memb_FirstName, Memb_LastName FROM Memb WITH (NOLOCK) WHERE Memb_No = " & vMembNo
      sOpenDb2
      Set oRs2 = oDb2.Execute(vSql)
      fMembName = oRs2("Memb_FirstName") & " " & oRs2("Memb_LastName")
      Set oRs2 = Nothing      
      sCloseDb2
      If Len(Trim(fMembName)) = 0 Then fMembName = " ... "
    End Function 
 

    Function spLogsTimeSpent (vMembNo, vProgId, vModsId, vStrDate, vEndDate)    
      Dim oRs
      spLogsTimeSpent = "0|0"
      sOpenDb5
      Set oCmd = Server.CreateObject("ADODB.Command")
      Set oCmd.ActiveConnection = oDb5
      oCmd.CommandType = adCmdStoredProc      
      With oCmd
        .CommandText = "spLogsTimeSpent"
        .Parameters.Append .CreateParameter("@MembNo",    adInteger, adParamInput,   , vMembNo)
        .Parameters.Append .CreateParameter("@ProgId",    adVarChar, adParamInput,  7, vProgId)
        .Parameters.Append .CreateParameter("@ModsId",    adVarChar, adParamInput, 11, vModsId)
        .Parameters.Append .CreateParameter("@StrDate",   adDate,    adParamInput,   , vStrDate)
        .Parameters.Append .CreateParameter("@EndDate",   adDate,    adParamInput,   , vEndDate)
      End With
      Set oRs = oCmd.Execute()
      If Not oRs.Eof Then 
        spLogsTimeSpent = oRs("Logs_No") & "|" & oRs("TimeSpent")
      End If  
      Set oRs = Nothing      
      Set oCmd = Nothing
      sCloseDb5
    End Function


    Function spLogsBestValues (vMembNo, vModsId, vStrDate, vEndDate)  
      Dim oRs
      sOpenDb5
      Set oCmd = Server.CreateObject("ADODB.Command")
      Set oCmd.ActiveConnection = oDb5
      oCmd.CommandType = adCmdStoredProc      
      With oCmd
        .CommandText = "spLogsBestValues"
        .Parameters.Append .CreateParameter("@MembNo",    adInteger, adParamInput,   , vMembNo)
        .Parameters.Append .CreateParameter("@ModsId",    adVarChar, adParamInput, 11, vModsId)
        .Parameters.Append .CreateParameter("@StrDate",   adDate,    adParamInput,   , vStrDate)
        .Parameters.Append .CreateParameter("@EndDate",   adDate,    adParamInput,   , vEndDate)
      End With
      Set oRs = oCmd.Execute()
      If oRs.Eof Then 
        spLogsBestValues = -1
      ElseIf IsNull(oRs("Logs_Grade")) Then
        spLogsBestValues = -1
      Else
        spLogsBestValues = Cint(oRs("Logs_Grade")) & "|" & fFormatSqlDate(oRs("Logs_Posted"))
      End If
      Set oRs = Nothing      
      Set oCmd = Nothing
      sCloseDb5

      '...if "completed" return "-999 | Date of first completion" if no score
      If spLogsBestValues <> -1 Then Exit Function
      
      sOpenDb5
      Set oCmd = Server.CreateObject("ADODB.Command")
      Set oCmd.ActiveConnection = oDb5
      oCmd.CommandType = adCmdStoredProc      
      With oCmd
        .CommandText = "spLogsCompleted"
        .Parameters.Append .CreateParameter("@MembNo",    adInteger, adParamInput,   , vMembNo)
        .Parameters.Append .CreateParameter("@ModsId",    adVarChar, adParamInput, 11, vModsId)
        .Parameters.Append .CreateParameter("@StrDate",   adDate,    adParamInput,   , vStrDate)
        .Parameters.Append .CreateParameter("@EndDate",   adDate,    adParamInput,   , vEndDate)
      End With
      Set oRs = oCmd.Execute()
      If Not oRs.Eof Then spLogsBestValues = "-999|" & fFormatSqlDate(oRs("Logs_Posted"))
      Set oRs = Nothing      
      Set oCmd = Nothing
      sCloseDb5
    End Function


    Function spLogsAttempts(vMembNo, vModsId, vStrDate, vEndDate)
      Dim oRs
      spLogsAttempts= "0"
      sOpenDb5
      Set oCmd = Server.CreateObject("ADODB.Command")
      Set oCmd.ActiveConnection = oDb5
      oCmd.CommandType = adCmdStoredProc      
      With oCmd
        .CommandText = "spLogsAttempts"
        .Parameters.Append .CreateParameter("@MembNo",    adInteger, adParamInput,   , vMembNo)
        .Parameters.Append .CreateParameter("@ModsId",    adVarChar, adParamInput, 11, vModsId)
        .Parameters.Append .CreateParameter("@StrDate",   adDate,    adParamInput,   , vStrDate)
        .Parameters.Append .CreateParameter("@EndDate",   adDate,    adParamInput,   , vEndDate)
      End With
      Set oRs = oCmd.Execute()
      If Not oRs.Eof Then 
        spLogsAttempts = oRs("Attempts")
      End If  
      Set oRs = Nothing      
      Set oCmd = Nothing
      sCloseDb5
    End Function
    

    Function spLogsBookmark (vMembNo, vModsId, vStrDate, vEndDate)    
      Dim oRs
      spLogsBookmark = "0|0"
      sOpenDb5
      Set oCmd = Server.CreateObject("ADODB.Command")
      Set oCmd.ActiveConnection = oDb5
      oCmd.CommandType = adCmdStoredProc      
      With oCmd
        .CommandText = "spLogsBookmark"
        .Parameters.Append .CreateParameter("@MembNo",    adInteger, adParamInput,   , vMembNo)
        .Parameters.Append .CreateParameter("@ModsId",    adVarChar, adParamInput, 11, vModsId)
        .Parameters.Append .CreateParameter("@StrDate",   adDate,    adParamInput,   , vStrDate)
        .Parameters.Append .CreateParameter("@EndDate",   adDate,    adParamInput,   , vEndDate)
      End With
      Set oRs = oCmd.Execute()
      If Not oRs.Eof Then 
        spLogsBookmark = oRs("Logs_No") & "|" & oRs("Bookmark")
      End If  
      Set oRs = Nothing      
      Set oCmd = Nothing
      sCloseDb5
    End Function
  %>


