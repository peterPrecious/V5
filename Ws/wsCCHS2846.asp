<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  If Request.Form("vFunction") <> "CCHS" Then
    Response.Write "error"

  Else
    Dim vFunction, vPassword, vId

    '...function returns one of:
    '   "password" meaning invalid password
    '   "ok"       meaning all tests passed ok
    '   "eof"      meaning that ID is not on file
    '   "error"    meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "CCHS2846"
    Session("MembId")     = "CCHS"
    Session("CustAcctId") = "2846"
    Session("HostDb")     = "V5_Vubz"

    vId = Ucase(Request.Form("vId"))
    sGetMembById Session("CustAcctId"), vId
    If vMemb_Eof Then 
      Response.Write "eof"
    Else
      Response.Write "ok"
    End If

    Session.Abandon 

  End If  
%>
