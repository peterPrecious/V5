<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<% Response.Redirect "/Chaccess/Signin" %>
<%
  '...if house account do not display, it will be reassigned at signin
  Dim vCust
  vCust = Request("vCust")
  If vCust = "VUBZ2294" Then vCust = ""  
%>  

<html>

<head>
  <title>:: Vubiz</title>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <script>
    function Validate(theForm) {
      if (theForm.vId.value == ""){
        var vMsg = "Incorporar por favor una contraseña válida.";
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
      <td width="100%" align="center"><h1>Conexión</h1><h2>Provea por favor su identificación de cliente <br>y contraseña, y entonces haga clic en <b>Conexión</b>.</h2>
      <table cellpadding="6" border="0" id="table12" style="border-collapse: collapse" bordercolor="#DDEEF9" width="175">
        <form method="GET" action="../Default.asp" target="_top"  onsubmit="return Validate(this)" name="fForm">
          <input type="hidden" name="vLang" value="ES">
          <tr>
            <td class="c2"><p class="c2">Identificación de cliente:<br><input type="text" name="vCust" size="14" value="<%=vCust%>"><br>Contraseña:<br><input type="password" name="vId" size="20" value="<%=Request("vId")%>"></p><p align="right"><input type="submit" value="Conexión" name="bGo" class="button"></p></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
  </table>

</body>

</html>