<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  Dim vMsg
  vMsg = fOkValue(Request("vMsg"))
  If Len(vMsg) > 0 Then vMsg = "(" & vMsg & ")" 
%>

<html>

<head>
  <title>EcomError</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <table class="table">
    <tr>
      <td style="text-align:center">
        <h1><br><!--[[-->Unfortunately there was an error in processing your e-commerce transaction.<!--]]--> </h1>
        <h2>
          <!--[[-->Please email details to<!--]]--> <a <%=fstatx%> href="mailto:info@vubiz.com?subject=Ecommerce error: <%=vMsg%>">info@vubiz.com</a>
          <br /><br /><%=vMsg%>
        </h2>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>
</html>