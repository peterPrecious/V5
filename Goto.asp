<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Urls_Routines.asp"-->

<%
  Dim vUrl

  '...get generated URLs from a Host DB table called Urls
  Session("Host")      = Lcase(Request.ServerVariables("HTTP_HOST") & "/V5")  : svHost      = Session("Host")
  Session("HostDbPwd") = "vudb2112mississauga"                                : svHostDbPwd = Session("HostDbPwd")
  Session("HostDb")    = "V5_Vubz"                                            : svHostDb    = Session("HostDb")

  vUrl = fGeturls(Request.QueryString("vCode"))

  Response.Redirect vUrl
%>