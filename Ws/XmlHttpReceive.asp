<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vParm, vResponse, oXmlHttp
  Dim vManagerId, vCommand, vAction

  '...extract mandatory parms
  vCust_Id            = Request.Form("vVubizAccount")
  vManagerId          = Request.Form("vManagerId") 
  vAction             = Request.Form("vAction")        '...not used, assume adding new members

  '...can be empty or full
  vMemb_Id            = Request.Form("vMemb_Id")

  '...extract optional parms
  vMemb_FirstName     = Request.Form("vMemb_FirstName") 
  vMemb_LastName      = Request.Form("vMemb_LastName")
  vMemb_Email         = Request.Form("vMemb_Email")

  '...extract learning program(s), ie P1001EN|P1202EN then change the pipes to spaces
  vMemb_Programs      = Request.Form("vMemb_Programs") 
  vMemb_Programs      = Replace(vMemb_Programs, "|", " ")

  '...expires can be empty, or no days, or date
  i                   = Request.Form("vMemb_Expires")
  If IsDate(i) Then
    vMemb_Expires = i
  ElseIf IsNumeric(i) Then
    If i > 0 And i < 367 Then
      vMemb_Expires = Now + i
      vMemb_Expires = DateAdd("d", i, Now)
    End If
  Else    
    vMemb_Expires = DateAdd("d", 90, Now)
  End If
  vMemb_Expires = fFormatDate(vMemb_Expires)


  If Len(vCust_Id) = 0 Or Len(vManagerId) = 0 Or Len(vAction) = 0 Then
    vResponse = "XmlHttp: Failure - Missing mandatory field(s)"
  End If

  If vResponse = "" Then sSignInOk (vCust_Id)

  If vResponse = "" Then 
    '...if not passing an member id then generate on
    If Len(vMemb_Id) = 0 Then
      sEnroll_Generate_Ok
    Else
      sEnroll_Existing_Ok
    End If
  End If

  If vResponse = "" Then vResponse = "XmlHttp: OK - " & vMemb_Id
  
  '...Return
  Response.Write vResponse


  '_______________________________________________________________________________

  Sub sSignInOk (vParm)
    Session("HostDb") = "V5_Vubz"      
    '...ensure customer is valid
    sGetCust (vCust_Id)
    If vCust_Eof Then 
      vResponse = "XmlHttp: Failure - Customer Account invalid"
      Exit Sub
    End If
    If vCust_ManagerId <> vManagerId Then
      vResponse = "XmlHttp: Failure - Manager Id is invalid"
      Exit Sub
    End If
    vMemb_AcctId = vCust_AcctId
  End Sub

  '...register/enroll the learner, must NOT be onfile
  Sub sEnroll_Existing_Ok
    sGetMembById vCust_AcctId, vMemb_Id
    If vMemb_Eof Then 
      sAddMemb
      If Not vFileOK Then
        vResponse = "XmlHttp: Failure - Unable to add learner"
        Exit Sub
      End If
    Else
      vResponse = "XmlHttp: Failure - Learner already on file"
      Exit Sub
    End If
  End Sub
  
  '...register/enroll the learner and return the Member Id
  Sub sEnroll_Generate_Ok
    '...get new password
    i = fNextMembNo
    vMemb_Id = Right("000000" & i, 6) & "-" & fSecurityCode (vCust_AcctId, i)
    sAddMemb
    If Not vFileOK Then
      vResponse = "XmlHttp: Failure - Unable to add learner"
      Exit Sub
    End If
  End Sub
%>