<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Arts.asp"-->
<% Session("HostDb") = "V5_Vubz"  '...set since bypassing "signin" %>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <style type="text/css">
    .cArticles {FONT-SIZE: 7.5pt; FONT-FAMILY: verdana}
    .cArticles A {TEXT-DECORATION: none}
    .cArticles A:hover {TEXT-DECORATION: underline}
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>
  <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <th><h1>Vu Articles</h1><p>Click to view any article of interest.</p></th>
    </tr>
    <%
    '...read Arts
    sOpenDb
    vSql = "Select * FROM Arts "
    Set oRs = oDb.Execute(vSQL)    
    Do While Not oRs.Eof 
      sReadArts
      If vArts_Type = "T" Then
    %>
    <tr>
      <td height="12"><p align="left"><br></p><font color="#3977B6" size="2"><b><%=vArts_Title%></b></font> </td>
    </tr>
    <%  
      ElseIf vArts_Type = "A" Then
    %>
    <tr>
      <td class="cArticles" height="12">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fstatx%> href="javascript:articles('<%=vArts_No%>')"><%=vArts_Title%></a>&nbsp; </td>
    </tr>
    <%  
      End If
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  %>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


