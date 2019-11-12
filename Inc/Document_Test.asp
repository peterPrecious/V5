<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Document.asp"-->

<head>
  <meta http-equiv="Content-Language" content="en-us">
  <script language="JavaScript" src="/V5/Inc/Launch.js"></script>
</head>

<%
Dim vFileName, vLang, vCust, vAcctId

vFileName   = fDefault(Request("vFileName"), "Affirmative.pdf") 
vLang       = fDefault(Request("vLang"), svLang)
vCust       = fDefault(Request("vCust"), Left(svCustId, 4)) 
vAcctId     = fDefault(Request("vAcctId"), svCustAcctId)

%>

<p><a href="#" onclick="fullScreen('//vubiz.com<%=fDocumentUrl(vFileName, "", vLang, vCust, vAcctId, "", "")%>')">Document</a></p>
