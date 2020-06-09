<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

<% 
  Server.Execute vShellHi 
  
  Dim vScript, vCertNameAvailable, vUrl

  '...get logos and see if custom
  sGetCust (svCustId)

  Session("CertSample")      = "y"

  '...Postback? 
  If Request.Form("vHidden") = "y" Then    
    
    Session("CertType")      = Request.Form("vCertType")
    Session("CertLang")      = Request.Form("vCertLang")
    Session("CertLogo1")     = Request.Form("vCertLogo1")
    Session("CertLogo2")     = Request.Form("vCertLogo2")
    Session("CertLogos")     = Request.Form("vCertLogos")
    Session("CertMark")      = Request.Form("vCertMark")
    Session("CertId")        = Request.Form("vCertId")
    Session("CertTitle")     = Request.Form("vCertTitle")
    Session("CertName")      = Request.Form("vCertName")

    vCertNameAvailable       = Request.Form("vCertNameAvailable")
    If vCertNameAvailable    = 0 Then
      Session("CertName")    = ""
    End If
  
'   Response.Redirect "../" & Session("CertLang") & "/Certificate.asp"

    '...if custom cert, run from the repository
    If vCust_CustomCert Then 
      vUrl = "../Repository/" & svHostDb & "/" & svCustAcctId & "/Tools/Certificate.asp"
    Else
      vUrl = "Certificate.asp"
    End If  

    vScript = ""
    vScript = vScript & "<script for='window' event='onload'>" & vbCrLf
    vScript = vScript & "  window.open('" & vUrl & "','Certificate','width=650,height=425,left=100,top=100,status=no,scrollbars=no,resizable=no')" & vbCrLf
    vScript = vScript & "</script>"
    
'   Response.Write Server.HtmlEncode(vScript)
    Response.Write vScript

  Else

    Session("CertType")      = "Exam"
    Session("CertLang")      = "EN"
    Session("CertLogo1")     = vCust_CertLogoVubiz
    Session("CertLogo2")     = vCust_CertLogoCust
    Session("CertLogos")     = "Both"
    vCertNameAvailable       = 1
    
  End If

%>

  <form method="POST" action="CertificateSample.asp">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="2" cellspacing="0">
      <tr>
        <td valign="top" colspan="2"><h1 align="center">Certificate Sample - Standard</h1><h2>This program shows how standard certificates will look under various circumstances.&nbsp; Note, this is for display only and does not actually configure the certificate.&nbsp; Configuration is an administrative feature and cannot be changed by customers.</h2></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Language :&nbsp;&nbsp; </th>
        <td valign="top"><input type="radio" name="vCertLang" value="EN" <%=fcheck(session("certlang"),"en")%>>EN&nbsp;&nbsp; <input type="radio" name="vCertLang" value="FR" <%=fcheck(session("certlang"),"fr")%>>FR&nbsp;&nbsp; <input type="radio" name="vCertLang" value="ES" <%=fcheck(session("certlang"),"es")%>>ES</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Certificate Type :&nbsp;&nbsp; </th>
        <td valign="top"><input type="radio" value="Test" name="vCertType" <%=fcheck(session("certtype"),"test")%>>Self Assessment (platform)<br><input type="radio" value="Exam" name="vCertType" <%=fcheck(session("certtype"),"exam")%> checked>Exam (platform)<br><input type="radio" value="Completion" name="vCertType" <%=fcheck(session("certtype"),"exam")%>>Completion (platform - when no test or exam)</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Host Logo :&nbsp;&nbsp; </th>
        <td valign="top">&nbsp;<input type="text" name="vCertLogo1" size="30" value="<%=Session("CertLogo1")%>"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Customer Logo :&nbsp;&nbsp; </th>
        <td valign="top">&nbsp;<input type="text" name="vCertLogo2" size="30" value="<%=Session("CertLogo2")%>"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Display Logo:&nbsp;&nbsp; </th>
        <td valign="top"><input type="radio" value="Vubiz" name="vCertLogos" <%=fcheck(session("certlogos"),"vubiz")%>>Host only<br><input type="radio" value="Cust" name="vCertLogos" <%=fcheck(session("certlogos"),"cust")%>>Customer only (CFIB, for example)<br><input type="radio" value="Both" name="vCertLogos" <%=fcheck(session("certlogos"),"both")%>>Both Host and Customer</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Learner Name available? :&nbsp;&nbsp; </th>
        <td valign="top"><input type="radio" name="vCertNameAvailable" value="1" <%=fcheck(vcertnameavailable,"1")%>>Yes<br><input type="radio" name="vCertNameAvailable" value="0" <%=fcheck(vcertnameavailable,"0")%>>No</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Sample Learner Name :&nbsp;&nbsp; </th>
        <td valign="top">&nbsp;<input type="text" name="vCertName" size="30" value="<%=Session("CertName")%>"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Assessment Name :&nbsp;&nbsp; </th>
        <td valign="top">&nbsp;<input type="text" name="vCertTitle" size="30" value="<%=Session("CertTitle")%>"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Assessment ID :&nbsp;&nbsp; </th>
        <td valign="top">&nbsp;<input type="text" name="vCertId" size="30" value="<%=Session("CertId")%>"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Mark / Completion :&nbsp;&nbsp; </th>
        <td valign="top">&nbsp;<input type="text" name="vCertMark" size="8" value=".85"> enter score as .85 and completion time as 85</td>
      </tr>
      <tr>
        <td align="center" valign="top" colspan="2">&nbsp;<p><input type="submit" value="Submit" name="bSubmit" class="button"></p><p>&nbsp;</p></td>
      </tr>
    </table>
    <input type="hidden" name="vHidden" value="y">
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
