<%
  Dim vMemb_AcctId, vMemb_Id, vMemb_No, vMemb_FirstName, vMemb_LastName, vMemb_Email 
  Dim vMemb_Eof, vMemb_Level, vMemb_Auth
  '____ Memb  ________________________________________________________________________

  '...Return Member List string of FirstName, LastName, No, Email
  Function fMemb_List(svCustAcctId)
    fMemb_List = ""
    vResponseNamePair = ""
    
    '...ignore inactive and only get authors
    vSql = "SELECT * FROM Memb WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Active = 1 AND Memb_Auth = 1 AND Memb_Id = '" & vMembID & "'"

   'sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    
    If Not oRs.Eof Then
      vMemb_Eof = False
      sReadMemb
      vResponseNamePair = vResponseNamePair _
       & "vMemb_No=" & vMemb_No _
       & "&vMemb_FirstName=" & vMemb_FirstName _
       & "&vMemb_LastName=" & vMemb_LastName _
       & "&vMemb_Level=" & vMemb_Level _
       & "&vMemb_Email=" & vMemb_Email & "&"
    Else  
      vMemb_Eof = True
    End If

    Set oRs = Nothing
    sCloseDb
    '...strip trailing "&"
    If Len(vResponseNamePair) > 3 Then vResponseNamePair = Left(vResponseNamePair, Len(vResponseNamePair)-1)

  End Function

  '...get the current fields from the current record in the record set
  Sub sReadMemb
    vMemb_AcctId      = oRs("Memb_AcctId")
    vMemb_Id          = oRs("Memb_Id")
    vMemb_No          = oRs("Memb_No")
    vMemb_FirstName   = oRs("Memb_FirstName")
    vMemb_LastName    = oRs("Memb_LastName")
    vMemb_Email       = oRs("Memb_Email")
    vMemb_Level       = oRs("Memb_Level")
    vMemb_Auth        = oRs("Memb_Auth")
  End Sub
  
  %>