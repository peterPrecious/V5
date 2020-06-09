<%
  '...reset from old LRC 
  Sub sAssessmentReset (vCustAcctId, vMembNo, vFirstName, vLastName, vModId) 

    '...confirm is something to reset (delete)
    vSql = " SELECT Logs_AcctId, Logs_MembNo, Left(Logs_Item, 6) AS [Mods_Id], Logs_Type " _
         & "   FROM Memb INNER JOIN Logs ON Memb.Memb_No = Logs.Logs_MembNo" _
         & "   WHERE Logs_AcctId = '" & vCustAcctId & "'" _
         & "     AND Logs_MembNo = " & vMembNo _ 
         & "     AND ISNULL(Memb_FirstName, '') = '" & fUnQuote(vFirstName) & "'" _
         & "     AND ISNULL(Memb_LastName, '') = '" & fUnQuote(vLastName) & "'" _
         & "     AND Left(Logs_Item, 6) = '" & vModId & "'" _
         & "     AND Logs_Type IN ('E','H','T','L','S')"
'   sDebug
    sOpenDb3
    Set oRs3 = oDb3.Execute (vSql)

'   Response.Write "<br>" & oRs3.Eof

    '...delete all session data except TimeSpent ('P')
    If Not oRs3.Eof Then 
      vSql = " DELETE Logs " _
           & "   WHERE Logs_AcctId = '" & vCustAcctId & "'" _
           & "     AND Logs_MembNo = " & vMembNo  _
           & "     AND Left(Logs_Item, 6) = '" & vModId & "'" _
           & "     AND Logs_Type IN ('E','H','T','L','S')"
'     sDebug
      oDb3.Execute (vSql)
    End If

    Set oRs3 = Nothing
    sCloseDb3
  
  End Sub

%>