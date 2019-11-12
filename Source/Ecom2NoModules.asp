<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
    Session("Ecom_Prog") = Request("vProgId")
    Session("Ecom_Mods") = ""

    sGetProg Session("Ecom_Prog")
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
      <td nowrap>
      <img border="0" src="../Images/Ecom/Modules.gif" width="75" height="67"></td>
      <td align="center">
      <h2><b><!--[[-->My Learning Modules<!--]]--></b></h2> <h6 align="center"><!--[[-->There are no modules available. <!--]]--></h6>
      </td>
    </tr>
  </table>



  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>

