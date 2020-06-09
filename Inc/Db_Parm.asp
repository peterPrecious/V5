<%
  Dim vParm_No, vParm_Value, vParm_Note

  '...Get Parm Record
  Function fParmValue (vParmNo)
    fParmValue = ""
    If Len(vParmNo) > 0 Then 
      If IsNumeric(vParmNo) Then 
        sOpenDb
        vSql = "SELECT Parm_Value FROM Parm WHERE (Parm_No = " & vParmNo & ")"
'       sDebug
        Set oRs = oDb.Execute(vSql)
        If Not oRs.Eof Then fParmValue = oRs("Parm_Value")
        Set oRs= Nothing
        sCloseDb
      End If
    End If
  End Function  

  Function fParmMods (vParmNo)
    fParmMods = ""
    If Len(vParmNo) > 0 Then 
      If IsNumeric(vParmNo) Then 
        sOpenDb
        vSql = "SELECT Parm_Mods FROM Parm WHERE (Parm_No = " & vParmNo & ")"
'       sDebug
        Set oRs = oDb.Execute(vSql)
        If Not oRs.Eof Then fParmMods = oRs("Parm_Mods")
        Set oRs= Nothing
        sCloseDb
      End If
    End If
  End Function

%>