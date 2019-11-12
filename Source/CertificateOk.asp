<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script for="window" event="onload">
  //  parent.frames.contents.location.href = parent.frames.contents.location.href;
  <% 
    Session("CertSample") = ""   '...ensure certificate is NOT a sample
    '...transfer to certificates
    Response.Write "  window.open('Certificate.asp','Certificate','toolbar=no,width=650,height=425,left=100,top=100,status=no,scrollbars=no,resizable=no')"
  %>
  </script>
</head>

<body link="#000080" vlink="#000080" alink="#000080" bgcolor="#FFFFFF" text="#000080">

  <table border="0" width="100%">
    <tr>
      <td width="100%" align="center"><h1><br>
      <!--[[-->Congratulations!<!--]]--></h1><h2>
      <!--[[-->Your Certificate, which is now displayed in a separate window,<br>can be printed by pressing &lt;Ctrl&gt;+P simultaneously.<!--]]--></h2>
      <h2>If you cannot see your certificate your browser may be configured to block pop-ups - please enable pop-ups while using the Vubiz service.</h2>
      </td>
    </tr>
  </table>

</body>

</html>

