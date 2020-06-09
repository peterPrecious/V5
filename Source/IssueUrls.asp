<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Urls.asp"-->

<%
  Session("HostDb") = "V5_Vubz"
  vUrls_Id = fNextUrlsId
  vUrls_Goto = "/V5/default.asp?vCust=IBMC9994&vId=egglestonton&vLang=EN"
  sInsertUrls  
%>