<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->

<head>
  <title>Certificate_Test</title>
  <meta http-equiv="Content-Language" content="en-us">
  <script src="/V5/Inc/Launch.js"></script>
</head>

<!--
           & "&vFirstName=" & fDefault(vFirstName, svMembFirstName) _
           & "&vLastName="  & fDefault(vLastName, svMembLastName) _
           & "&vScore="     & vScore _
           & "&vDate="      & fDefault(vDate, fFormatDate(Now)) _
           & "&vModsId="    & vModsId _
           & "&vTitle="     & vTitle _
           & "&vLang="      & fDefault(vLang, svLang) _
           & "&vCust="      & Left(fDefault(vCust, svCustId), 4) _
           & "&vAcctId="    & fDefault(vCust, svCustAcctId) _
           & "&vProgId="    & vProgId _
           & "&vLogo="      & fDefault(vLogo, svCustBanner) _
           & "&vMemo="      & vMemo

<p><a href="#" onclick="fullScreen('<%=fCertificateUrl("P�ter", "Bulloch", 80, "Jan 10, 2010", "1234EN", "Test Assessment", "FR", "VUBZ", "2274", "P1234EN", "vubz.jpg", "Eat My Shorts", "")%>')">Certificate</a></p>
<p><a href="#" onclick="fullScreen('<%=fCertificateUrl("P�ter", "Bulloch", 80, "2010/1/10", "1234EN", "Test Assessment", "FR", "VUBZ", "2274", "P1234EN", "vubz.jpg", "Eat My Shorts", "")%>')">Certificate</a></p>

-->


<p><a href="#" onclick="fullScreen('<%=fCertificateUrl("P�ter", "Bulloch", 80, "", "1234EN", "Test Assessment", "FR", "VUBZ", "2274", "P1234EN", "vubz.jpg", "Eat My Shorts", "1234567")%>')">Certificate</a></p>
