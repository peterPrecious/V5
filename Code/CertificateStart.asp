<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
    </head>

  <body>

<% Server.Execute vShellHi %>

<%
  Dim vScript, vCertNameAvailable, vUrl

  '...get logos and see if custom
  sGetCust (svCustId)

  Session("CertID")          = "1234EN"
  Session("CertMark")        = .85
  Session("CertTitle")       = "Sample Module/Program Title"
  Session("CertName")        = "Jean Learner"
  Session("CertSample")      = "y"

  '...Postback? 
  If Request.Form("vHidden") = "y" Then    
    
    Session("CertType")      = Request.Form("vCertType")
    Session("CertLang")      = Request.Form("vCertLang")
    Session("CertLogo1")     = Request.Form("vCertLogo1")
    Session("CertLogo2")     = Request.Form("vCertLogo2")
    Session("CertLogos")     = Request.Form("vCertLogos")
    vCertNameAvailable       = Request.Form("vCertNameAvailable")

    If Session("CertType")   = "Completion" Then
      Session("CertMark")    = 85
    Else
      Session("CertMark")    = .85
    End If

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

<form method="POST" action="CertificateStart.asp">
  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3" cellspacing="0">
    <tr>
      <td valign="top" colspan="2">
      <h1 align="center">Certificate Sample</h1>
      <h2>This program shows how certificates will look under various circumstances.&nbsp; Note, this is for display only and does not actually configure the certificate.&nbsp; Configuration is an administrative feature and cannot be changed by customers.<br>&nbsp;</h2>
      </td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Language :&nbsp;&nbsp; </th>
      <td valign="top">
          <input type="radio" name="vCertLang" value="EN" <%=fCheck(Session("CertLang"),"EN")%>>EN&nbsp;&nbsp; 
          <input type="radio" name="vCertLang" value="FR" <%=fCheck(Session("CertLang"),"FR")%>>FR&nbsp;&nbsp; 
          <input type="radio" name="vCertLang" value="ES" <%=fCheck(Session("CertLang"),"ES")%>>ES
      </td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Certificate Type :&nbsp;&nbsp; </th>
      <td valign="top"><input type="radio" value="Test" name="vCertType" <%=fCheck(Session("CertType"),"Test")%>>Test<br><input type="radio" value="Exam" name="vCertType" <%=fCheck(Session("CertType"),"Exam")%> checked>Exam<br><input type="radio" value="Completion" name="vCertType" <%=fCheck(Session("CertType"),"Exam")%>>Completion (when no test or exam)<br>&nbsp;</td>
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
      <td valign="top"><input type="radio" value="Vubiz" name="vCertLogos" <%=fCheck(Session("CertLogos"),"Vubiz")%>>Host only<br><input type="radio" value="Cust" name="vCertLogos" <%=fCheck(Session("CertLogos"),"Cust")%>>Customer only (CFIB, for example)<br><input type="radio" value="Both" name="vCertLogos" <%=fCheck(Session("CertLogos"),"Both")%>>Both Host and Customer</td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Learner Name available? :&nbsp;&nbsp; </th>
      <td valign="top"><input type="radio" name="vCertNameAvailable" value="1" <%=fCheck(vCertNameAvailable,"1")%>>Yes<br><input type="radio" name="vCertNameAvailable" value="0" <%=fCheck(vCertNameAvailable,"0")%>>No</td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Sample Learner Name :&nbsp;&nbsp; </th>
      <td valign="top">&nbsp;<%=Session("CertName")%></td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Exam/Test Name :&nbsp;&nbsp; </th>
      <td valign="top">&nbsp;<%=Session("CertTitle")%></td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Exam/Test No :&nbsp;&nbsp; </th>
      <td valign="top">&nbsp;<%=Session("CertID")%></td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap>Mark/Completion :&nbsp;&nbsp; </th>
      <td valign="top">&nbsp;85% or 85 minutes</td>
    </tr>
    <tr>
      <td align="center" valign="top" colspan="2"><br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" name="I1" type="image"><br>&nbsp;</td>
    </tr>
  </table>
  <input type="hidden" name="vHidden" value="y">
</form>

<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body></html>

