<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  If Ucase(Request.Form("vFunction")) <> "BCTF" Then

    Response.Write "error"

  Else

    Dim vFunction, vPassword, vCust, vId

    '...function returns one of:
    '   "password" meaning invalid password
    '   "ok"       meaning all tests passed ok
    '   "eof"      meaning that ID is not on file
    '   "error"    meaning invalid function call

    vPassword = Ucase(Request.Form("vPassword"))
    vCust     = Ucase(Request.Form("vCust"))
    vId       = Ucase(Request.Form("vId"))

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = vCust
    Session("MembId")     = "DRIVER"
    Session("CustAcctId") = Right(vCust, 4)
    Session("HostDb")     = "V5_Vubz"


    '...see if password is valid 
    If vPassword <> "BCTF" Then
      Response.Write "password"

    '...see if ID is on file
    Else

      sGetMembById Session("CustAcctId"), vId
      If vMemb_Eof Then 
        Response.Write "eof"
      Else
        Response.Write "ok"
      End If
    End If

    Session.Abandon 

  End If  
%>
