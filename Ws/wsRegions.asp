<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  If Request.Form("vFunction") <> "Regions" Then

    Response.Write "error"

  Else

    Dim vFunction, vPassword, vId

    '...function returns one of:
    '   "password" meaning invalid password
    '   "ok"       meaning all tests passed ok
    '   "eof"      meaning that ID is not on file
    '   "error"    meaning invalid function call

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = "RGMG2730"
    Session("MembId")     = "REGIONS"
    Session("CustAcctId") = "2730"
    Session("HostDb")     = "V5_Vubz"

    vPassword = Request.Form("vPassword")
    vId       = Request.Form("vId")

    '...see if password is valid 
    If Ucase(vPassword) <> "REGIONS1" Then
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
