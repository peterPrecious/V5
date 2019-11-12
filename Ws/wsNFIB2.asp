<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  If Request.Form("vFunction") <> "NFIB" Then

    Response.Write "error"

  Else

    Dim vFunction, vPassword, vId

    '...function returns one of:
    '   "password" meaning invalid password
    '   "ok"       meaning all tests passed ok
    '   "eof"      meaning that ID is not on file
    '   "error"    meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "NFIB2813"
    Session("MembId")     = "NFIB"
    Session("CustAcctId") = "2813"
    Session("HostDb")     = "V5_Vubz"

    vId                   = Ucase(Request.Form("vId"))

    sGetMembById Session("CustAcctId"), vId
    If vMemb_Eof Then 
      Response.Write "eof"
    Else
      Response.Write "ok"
    End If

    Session.Abandon 

  End If  
%>
