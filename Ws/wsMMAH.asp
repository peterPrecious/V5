<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vFunction, vCust, vId 

  '...determine if the user ID and Password are valid when signing in
  If Request.Form("vFunction") = "MMAH_SignIn" Then

    '...function returns one of:
    '   "invalid id"    meaning invalid user Id
    '   "ok"            meaning all tests passed ok
    '   "error"         meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "MMAH2782"
    Session("MembId")     = "WebService"
    Session("CustAcctId") = "2782"
    Session("HostDb")     = "V5_Vubz"


    '...see if Cust is valid
    vCust = Ucase(Request.Form("vCust"))

    '...see if ID is on file
    vId  = Ucase(Request.Form("vId"))
    sGetMembById Session("CustAcctId"), vId

    If vMemb_Eof Then 
      Response.Write "invalid id"
    Else
      Response.Write "ok"
    End If
    Session.Abandon 



  '...determine if the user ID and Password are valid when enrolling
  ElseIf Request.Form("vFunction") = "MMAH_Enroll" Then

    '...function returns one of:
    '   "invalid id"    meaning invalid user Id
    '   "ok"            meaning all tests passed ok
    '   "error"         meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "MMAH2782"
    Session("MembId")     = "WebService"
    Session("CustAcctId") = "2782"
    Session("HostDb")     = "V5_Vubz"

    '...see if ID is already on file
    vId  = Ucase(Request.Form("vId"))
    sGetMembById Session("CustAcctId"), vId

    If Not vMemb_Eof Then 
      Response.Write "invalid id"
    Else
      Response.Write "ok"
    End If
    Session.Abandon 

  Else
    Response.Write "error"

  End If  
  
%>
