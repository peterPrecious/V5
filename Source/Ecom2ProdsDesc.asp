<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prod.asp"-->
<% 
  sGetProd Request("vProdId")
%>
<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <base target="_self">
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellpadding="3" style="border-collapse: collapse">
    <tr>
      <td nowrap><img border="0" src="../Images/Ecom/Book.jpg" width="80" height="73"></td>
      <td width="90%">
      <h2 align="center"><b>Item Description</b></h2>
      <h2 align="center"><a <%=fstatx%> href="javascript:history.back(-1)">Return to Product List </a></h2>
      </td>
    </tr>
  </table>
  <table cellspacing="0" border="1" id="table3" width="100%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td>
      <h1><%=vProd_Title%></h1>
      <h2><%=vProd_Desc%></h2>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
