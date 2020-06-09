<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<% 
  '...this routine accepts transactions from CCHS OSHworks - essentially adding Programs to Learners and removing stati (3 services)

  Dim vCustId, vAcctId, vId, vFirstName, vLastName, vEmail, vPrograms, vStatus, vErrMsg

  '...is this a test?
  bTest = fIf(Lcase(Request("vTest")) = "y", True, False)
  '...is this a validate or commit?
  bCommit = fIf(Lcase(Request("vAction")) = "v", False, True)


  '...valid WS request?
  If Request("WS") <> "02" Then
    vStatus = 490
  Else
    vStatus         = 0 '... status not yet set
    vCustId         = Request("vCustId")
    vAcctId         = Right(Request("vCustId"), 4)    
    vId             = Ucase(Trim(Request.Form("vId")))
    vPrograms       = Trim(Ucase(Trim(Request.Form("vPrograms"))))
    If Not fCustG2Ok(vCustId) Then 
      vStatus = 401
    ElseIf Len(vPrograms) = 0 Then 
      vStatus = 402
    End If
    If vStatus = 0 Then
      sMemb_Empty     '...clean out any memb variables
      vMemb_AcctId    = vAcctId
      vMemb_Id        = vId
      sGetMembById vMemb_AcctId, vMemb_Id
      vStatus = fG2AssignNo(vAcctId, vPrograms, vMemb_Programs) 
      If vStatus = 0 Then 
        vMemb_FirstName = fUnquote(Request.Form("vFirstName"))
        vMemb_LastName  = fUnquote(Request.Form("vLastName"))
        vMemb_Email     = Trim(fUnquote(Request.Form("vEmail")))
        vMemb_Programs  = vMemb_Programs & " " & vPrograms  
        vMemb_ProgramsAdded = Now()               
        sAddMemb vMemb_AcctId
        vStatus = 202
      End If                
    End If
  End If

  sRaiseError vStatus




  Sub sRaiseError (vStatus)
    Dim vMsg
    Select Case vStatus
      Case 201  : vMsg = vStatus & " Successful Request" 
      Case 202  : vMsg = vStatus & " Successful Post" 
      Case 203  : vMsg = vStatus & " Successful Request" 
      Case 401  : vMsg = vStatus & " Customer ID is not a valid G2 Account" 
      Case 402  : vMsg = vStatus & " No Programs Assigned" 
      Case 403  : vMsg = vStatus & " Assigning a Program that is already Assigned" 
      Case 404  : vMsg = vStatus & " Assigning a Program that is not available" 
      Case 490  : vMsg = vStatus & " Invalid Web Service Request" 
    End Select    
    If bTest Then
      Response.Write vMsg & "<br>"
    Else
      Response.Status = vMsg
    End If
  End Sub


  Function fG2AssignNo (vAcctId, vProgsNew, vProgsOld) 
    Dim vProgsAss, aNew, aOld, aAss, i, j, k
    fG2AssignNo = 0
    '... first ensure that we are not getting an ID that has already been assigned
    aNew = Split(vProgsNew)
    aOld = Split(vProgsOld)
    For i = 0 To Ubound(aNew)
      For j = 0 To Ubound(aOld)
        If aOld(j) = aNew(i) Then
          fG2AssignNo = 403 : Exit Function
        End If
      Next
    Next
    '... next grab all available IDs
    sOpenCmd
    With oCmd
      .CommandText = "spG2Progs"
      .Parameters.Append .CreateParameter("@AcctId",  adVarChar,  adParamInput,    4, vAcctId)
    End With
    oCmd.Execute()
    Set oRs = oCmd.Execute()
    Do While Not oRs.Eof
      If oRs("Available") > 0 Then
        vProgsAss = vProgsAss & " " & oRs("ProgramId")
      End If
      oRs.MoveNext
    Loop                
    Set oRs = Nothing      
    Set oCmd = Nothing
    sCloseDb
    '... and ensure the new Prog(s) are available
    aNew = Split(vProgsNew)
    For i = 0 To Ubound(aNew)
      If Instr(vProgsAss, aNew(i)) = 0 Then
        fG2AssignNo = 404 : Exit Function
      End If
    Next
  End Function


%>

