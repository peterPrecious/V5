<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Exam_Routines.asp"-->
<html>

<head>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <title>Exam Manipulation</title>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <base target="_self">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <div align="center">
    <center>
    <table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td align="center"><h1>
        <!--[[-->Examination Instructions<!--]]--></h1>
        <h6><br>
        <!--[[-->You have manipulated the Browser in such a way that the current test has been contaminated<!--]]-->.<br><br>
        <!--[[-->For future reference, ONLY click buttons available to you within the Examination pages<!--]]-->.<br><br>&nbsp;</h6>
        <p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a> <br><br>&nbsp;</p></td>
      </tr>
    </table>
    </center></div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


