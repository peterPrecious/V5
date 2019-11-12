<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/DocumentKMX.asp"-->

<%
  '...similiar to Certificate.asp/Document.asp - variant for KMX
  Dim vFileName, vCustId, vLang
  
' //vubiz.com/v5/document.asp?vCustId=VUBZ2277&vFileName=harassment.pdf&vMembNo=1034754&vProgId=P2362EN&vModsId=1630EN&vLang=EN
' //vubiz.com/v5/documentKMX.asp?vCustId=21952&vFileName=harrasment.pdf&vLang=EN

  vFileName   = Request("vFileName")
  vCustId     = Request("vCustId")
  vLang       = Request("vLang")

  Response.Redirect fDocumentUrl(vFileName, "", vLang, vCustId, "", "", "")
%>