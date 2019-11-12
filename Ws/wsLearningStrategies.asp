<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vParm, vResponse, vPassword, oXmlHttp
  Dim vManagerId, vCommand, vAction

  vPassword           = Request.Form("vPassword")
  vCust_Id            = Request.Form("vCust_Id")
  vMemb_Id            = "" '...always setup new user
  vMemb_FirstName     = fUnquote(Request.Form("vMemb_FirstName")) 
  vMemb_LastName      = fUnQuote(Request.Form("vMemb_LastName"))
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


  If vPassword <> "1010101" Then
    vResponse = "XmlHttp: Failure - Invalid Password"
  ElseIf Len(vCust_Id) <> 8 Then
    vResponse = "XmlHttp: Failure - Invalid Customer Id"
  Else
    sSignInOk (vCust_Id)
    sEnroll_Generate_Ok
    vResponse = "XmlHttp: OK - " & vMemb_Id
  End If
  
  
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

  '...register/enroll the learner and return the Member Id
  Sub sEnroll_Generate_Ok
    '...get new password
    vMemb_AcctId = vCust_AcctId
    i = fNextMembNo
    vMemb_Id = Right("000000" & i, 6) & "-" & fSecurityCode (vMemb_AcctId, i)
    sAddMemb
    If Not vFileOK Then
      vResponse = "XmlHttp: Failure - Unable to add learner"
      Exit Sub
    End If
  End Sub
%>