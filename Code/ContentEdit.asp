<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim vProgram, aProgram
  vProgram = Request("vProgram")
  vProgram = vProgram & "~~~~~~~~~~~~~~~~~~~~~~~~~" '...add to ensure all values are received since this was written before the new parms were added
  aProgram = Split(vProgram,"~")
  i = -1 '...use as a starter for field values
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body text="#000080" vlink="#000080" alink="#000080" link="#000080" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <% Server.Execute vShellHi %>
  <table style="BORDER-COLLAPSE: collapse" bordercolor="#DDEEF9" border="1" width="100%">
    <tr>
      <td valign="top" colspan="2" bgcolor="#DDEEF9"><b><font face="Verdana" size="1">Program/Course Details</font></b></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Program :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Online $US :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Online $CA :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Program Length Hours :</font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Program Length Days :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">CDs Available ?&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">VuBooks Available ?&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Include Self Assessments :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Exam Id :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Min Questions :&nbsp; </font></b></td>
      <td><% i = i + 1%><%=aProgram(i)%>&nbsp; </td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Max Attempts :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Time Limit per Bank :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
    <tr>
      <td align="right"><b><font face="Verdana" size="1">Passing Grade :&nbsp; </font></b></td>
      <td><font face="Verdana" size="1">&nbsp;<% i = i + 1%><%=aProgram(i)%></font></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

