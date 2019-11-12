<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<% Response.Redirect "/Chaccess/Signin" %>
<%
  '...if house account do not display, it will be reassigned at signin
  Dim vCust
  vCust = Request("vCust")
  If vCust = "VUBZ2275" Then vCust = ""  
%>  

<html>

<head>
  <title>:: Vubiz</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <script>
    function Validate(theForm) {
      if (theForm.vId.value == ""){
        var vMsg = "Veuillez entrer un mot de passe valide.";
        alert(vMsg);
        theForm.vId.focus();
        return (false);
      }
      return (true);
    }
  </script>
  <base target="_self">
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <table cellpadding="10" width="100%" border="0" cellspacing="0" id="table11">
    <tr>
      <td width="100%" align="center"><h1>Inscription</h1><h2>Veuillez écrire votre identification de client et le mot de passe <br>puis cliquez <b>Inscription</b></h2>
      <table cellspacing="0" cellpadding="6" border="0" id="table12" bordercolor="#DDEEF9" width="175">
        <form method="GET" action="../Default.asp" target="_top"  onsubmit="return Validate(this)" name="fForm">
          <input type="hidden" name="vLang" value="FR">
          <tr>
            <td class="c2"><p class="c2">Identification de client :<br><input type="text" name="vCust" size="14" value="<%=vCust%>"><br>Mot de passe :<br><input type="password" name="vId" size="20" value="<%=Request("vId")%>"></p><p align="right"><input type="submit" value="Inscription" name="bGo" class="button"></p></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
  </table>

</body>

</html>
