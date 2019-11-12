<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vFunction, vCust, vId 

  '...get Cust Id and Learner ID
  vCust = Ucase(Request.Form("vCust"))
  vId   = Ucase(Request.Form("vId"))

  '...set these to ensure you can access the DB without signing in 
  Session("CustId")     = vCust
  Session("MembId")     = "WebService"
  Session("CustAcctId") = Right(vCust, 4)
  Session("HostDb")     = "V5_Vubz"

  sGetMembById Session("CustAcctId"), vId

  '...function returns one of:
  '   "invalid id"    meaning invalid user Id
  '   "ok"            meaning all tests passed ok
  '   "error"         meaning invalid function call


  '...determine if the user ID and Password are valid when signing in
  If Request.Form("vFunction") = "SignIn" Then
    If vMemb_Eof Then 
      Response.Write "invalid id"
    Else
      Response.Write "ok"
    End If

  '...determine if the user ID and Password are valid when enrolling
  ElseIf Request.Form("vFunction") = "Enroll" Then
    If Not vMemb_Eof Then 
      Response.Write "invalid id"
    Else
      Response.Write "ok"
    End If

  Else
    Response.Write "error"

  End If  

  Session.Abandon 
 
%>
