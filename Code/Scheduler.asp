<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<% 
  Dim vUrl
  sGetCust svCustId
  vUrl = "/gold/vuscheduler/default.aspx?perm=30720&vCust_No=" & vCust_No & "&vLang=" & svLang
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body style="overflow: hidden;" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">


  <div id="divSchedule" style="position: absolute; left:5%; top:5%; width:90%; height:90%;">
    <iframe id="iSchedule" name="iSchedule" style="background-color:#FFFFFF" border="0" frameborder="0" src="<%=vUrl%>" style="background-color: #FFFFFF" marginwidth="1" marginheight="1" height="100%" width="100%"></iframe>
  </div>


</body>

</html>





