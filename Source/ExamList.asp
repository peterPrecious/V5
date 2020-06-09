<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<html>
  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  </head>

  <body>

  <% Server.Execute vShellHi %>

  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td colspan="2">
      <h1>Exam Listing</h1>
      <%
      If Len(Request.QueryString("VMess")) > 0 Then
        Response.Write "<font color=red>" & Request.QueryString("VMess") & "</font>"
      End If
      %>
      </td>
    </tr>
    <tr>
      <th nowrap bgcolor="#DDEEF9" align="left" height="20">Exam Id</th>
      <th nowrap bgcolor="#DDEEF9" align="left" height="20">Title</th>
    </tr>
    <%
      '...read Mod info
      Dim vModId, vTitle
      sOpenDbBase
      vSql = "Select * FROM TstH "
      Set oRsBase = oDbBase.Execute(vSQL)    
      Do While Not oRsBase.EOF 
        vModId = oRsBase("TstH_Id")
        vTitle = oRsBase("TstH_Title")
    %>
    <tr>
      <td><a href="ExamView.asp?vModId=<%=vModId%>"><%=vModId%></a>&nbsp; </td>
      <td><%=vTitle%>&nbsp; </td>
    </tr>
    <%  
        oRsBase.MoveNext
      Loop
      Set oRsBase = Nothing
      sCloseDbBase    
    %>
  </table>
  <p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a></p>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>
