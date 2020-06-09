<%
  Dim vDial_No, vDial_AcctId, vDial_TskHNo, vDial_Posted, vDial_Message, vDial_MembNo, vDial_Type
  Dim vDial_Eof
  
  '____ Dial  ________________________________________________________________________

  '...get a message using the message no  
  Sub sGetDial (vNo)
    If fNoValue(vDial_Type) Then vDial_Type = 0
    vSql = "SELECT * FROM (Dial LEFT JOIN Memb ON Dial_MembNo = Memb.Memb_No) WHERE Dial_No = " & vNo & " AND Dial_Type = '" & vDial_Type & "'"
    sOpenDb
'   sDebug
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      vDial_Eof = 0
      sReadDial
    Else
      vDial_Eof = 1
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub

  '...get a message set
  Sub sGetDial_Rs (vType)
    If vType = "N" Then
      vSql = "SELECT * FROM Dial LEFT JOIN Memb ON Dial_MembNo = Memb_No WHERE Dial_AcctId = '" & svCustAcctId & "' AND Dial_TskHNo = " & vTskH_No  & " AND Dial_Type = '" & vType & "' AND Dial_MembNo = " & svMembNo & " ORDER BY Dial_Posted DESC"
    Else
      vSql = "SELECT * FROM Dial LEFT JOIN Memb ON Dial_MembNo = Memb_No WHERE Dial_AcctId = '" & svCustAcctId & "' AND Dial_TskHNo = " & vTskH_No  & " AND Dial_Type = '" & vType & "' ORDER BY Dial_Posted DESC"
    End If

'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
  End Sub

  '...get last message 
  Function fDial_LastPosted
    fDial_LastPosted = ""
    vSql = "SELECT TOP 1 Dial_Posted FROM Dial WHERE Dial_AcctId = '" & svCustAcctId & "' AND Dial_TskHNo = " & vTskH_No  & " AND Dial_MembNo <> " & svMembNo & " ORDER BY Dial_Posted DESC"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then Exit Function
    fDial_LastPosted = oRs2("Dial_Posted")
    sCloseDb2    
  End Function

  '...get the current fields from the current record in the record set
  Sub sReadDial
    vDial_No        = oRs("Dial_No")
    vDial_MembNo    = oRs("Dial_MembNo")
    vDial_Posted    = oRs("Dial_Posted")
    vDial_Message   = oRs("Dial_Message")
    vDial_Type      = oRs("Dial_Type")
    sReadMemb
  End Sub
  
  Sub sInsertDial
    If fNoValue(vDial_Type) Then vDial_Type = 0
    vSql = "INSERT INTO Dial "
    vSql = vSql & "(Dial_AcctId, Dial_TskHNo, Dial_Message, Dial_MembNo, Dial_Type)"
    vSql = vSql & " VALUES ('" & svCustAcctId & "', " & vTskH_No & ", '" & fUnQuote(vDial_Message) & "', " & svMembNo & ", '" & vDial_Type & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sUpdateDial
    vSql = "UPDATE Dial SET"
    vSql = vSql & " Dial_Message = '" & fUnquote(vDial_Message) & "' " 
    vSql = vSql & " WHERE Dial_No = " & vDial_No
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...delete Dial
  Sub sDeleteDial
    sOpenDb
    vSql = "DELETE FROM Dial WHERE Dial_No = " & vDial_No
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


%>