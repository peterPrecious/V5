<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->


<html>
 <head>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
   <title>:: Server Variables</title>
 </head>


  <body>

  <table>
    <tr>
      <td><b>Server Variable</b></td>
      <td><b>Value</b></td>
    </tr>
    <% For Each vFld In Request.ServerVariables %>
    <tr>
      <td><%= vFld %> </td>
      <td><%= Request.ServerVariables(vFld) %> </td>
    </tr>
    <% Next %>
  </table>


  </body>
</html>
