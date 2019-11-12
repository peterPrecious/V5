<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<html>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  </head>

  <body>

    <% Server.Execute vShellHi %>

    <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <th bgcolor="#DDEEF9" nowrap>Exam Id</th>
        <th bgcolor="#DDEEF9" align="left">Title</th>
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
        <td align="center" valign="top"><%=vModId%></td>
        <td valign="top"><%=vTitle%></td>
      </tr>
      <%  
          oRsBase.MoveNext
        Loop
        Set oRsBase = Nothing
        sCloseDbBase    
      %>
    </table>

    <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  </body>
</html>


