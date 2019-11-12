<!--#include virtual = "V5/Inc/Setup.asp"-->

<head>
  <base target="_self">
</head>

<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<% 
  sGetQueryString
  Session("Lang")   = vLang '...this session variable required by public site
  Session("CustId") = vCust '...this session variable required by public tab
  svCustId = vCust
  sGetCust svCustId
  Response.Redirect "/V5/Code/Ecom2Start.asp?vClose=Y&vMode=More&vContentOptions=" & vContentOptions
%>

