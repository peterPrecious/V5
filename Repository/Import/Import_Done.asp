<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vUrl1, vUrl2
  vUrl1 = "Default.asp?vCust=" & svCustId & "&vId=" & svMembId
  vUrl2 = "http://" & Replace(svDomain, "ww2.", "") & "/v5/Default.asp?vCust=" & svCustId & "&vId=" & svMembId & "&vGoto=Default.asp~3vPage~2UsersOK.asp"
  Session.Abandon 
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="http://vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <title>Upload</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <%  Server.Execute vShellHi %> 

  <center>
  <h2>Thank you, this session has successfully terminated.</h2>
  <a class="c2" href="<%=vUrl1%>">Click here to re start Import.</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a class="c2" href="<%=vUrl2%>">Click here to see User report.</a> </center>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>