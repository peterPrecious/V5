<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<html>

<head>
  <title>Generate Certificate</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
  Server.Execute vShellHi 

  Dim vLangId, vMembId, vModsId  

  sGetCust (svCustId)

  If Request.Form.Count > 0 Then
    vLangId = Request.Form("vLang_Id")
    vMembId = Request.Form("vMemb_Id")
    vModsId = Request.Form("vMost_Id")
  End If


  %>

  <form method="POST" action="CertificateSample.asp">
    <table class="table">
      <tr>
        <td colspan="2">
          <h1>Generate a Certificate</h1>
          <h2>This will generate a certificate for the Learner Id using the best score available for the Module Id. <br /><br /></h2>
        </td>
      </tr>
      <tr>
        <th>Language :</th>
        <td>
          <input type="radio" name="vLangId" value="EN" <%=fcheck(vLangId,"en")%>>EN<br />
          <input type="radio" name="vLangId" value="FR" <%=fcheck(vLangId,"fr")%>>FR<br />
          <input type="radio" name="vLangId" value="ES" <%=fcheck(vLangId,"es")%>>ES
        </td>
      </tr>
      <tr>
        <th>Learner Id :</th>
        <td>
          <input type="text" name="vMembId" size="30" value="<%=vMembId%>"></td>
      </tr>
      <tr>
        <th>Module Id :</th>
        <td>
          <input type="text" name="vModsId" size="30" value="<%=vModsId%>"></td>
      </tr>

      <tr>
        <td colspan="2">
          <input type="submit" value="Submit" name="bSubmit" class="button">
        </td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


