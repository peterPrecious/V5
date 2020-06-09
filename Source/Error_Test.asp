<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %> 
  <div align="center">
    <table border="0" width="60%" id="table1" cellspacing="0" cellpadding="10">
      <tr>
        <td align="center"> 
          Test<p>

          <a href="error.asp?verr=poop&vreturn=n">message</a></p><p>
          <a href="error.asp?verr=poop">message, return here</a></p><p>
          <a href="error.asp?verr=poop&vreturn=//cnn.com">message, go to cnn.com</a>

          </p><p>&nbsp;
        </td>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->



</body>

</html>
