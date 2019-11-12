<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vFunction, vCust, vId 

  '...determine if the user ID and Password are valid when signing in
  If Request.Form("vFunction") = "WorkSmart_SignIn" Then

    '...function returns one of:
    '   "invalid cust"  meaning invalid cust (campus id)
    '   "invalid id"    meaning invalid user Id
    '   "ok"            meaning all tests passed ok
    '   "error"         meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "MEVT2747"
    Session("MembId")     = "WORKSMART"
    Session("CustAcctId") = "2747"
    Session("HostDb")     = "V5_Vubz"

    '...see if Cust is valid
    vCust = Ucase(Request.Form("vCust"))

    '...see if ID is on file
    vId  = Ucase(Request.Form("vId"))
    sGetMembById Session("CustAcctId"), vId

    If vCust <> "MEVT2747" Then 
      Response.Write "invalid cust"
    ElseIf vMemb_Eof Then 
      Response.Write "invalid id"
    Else
      Response.Write "ok"
    End If
    Session.Abandon 



  '...determine if the user ID and Password are valid when enrolling
  ElseIf Request.Form("vFunction") = "WorkSmart_Enroll" Then

    '...function returns one of:
    '   "invalid cust"  meaning invalid cust (campus id)
    '   "invalid id"    meaning invalid user Id
    '   "ok"            meaning all tests passed ok
    '   "error"         meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "MEVT2747"
    Session("MembId")     = "WORKSMART"
    Session("CustAcctId") = "2747"
    Session("HostDb")     = "V5_Vubz"

    '...see if Cust is valid
    vCust = Ucase(Request.Form("vCust"))

    '...see if ID is already on file
    vId  = Ucase(Request.Form("vId"))
    sGetMembById Session("CustAcctId"), vId

    If vCust <> "MEVT2747" Then 
      Response.Write "invalid cust"
    ElseIf Not vMemb_Eof Then 
      Response.Write "invalid id"
    Else
      Response.Write "ok"
    End If
    Session.Abandon 

  Else
    Response.Write "error"

  End If  
  
%>
