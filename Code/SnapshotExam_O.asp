<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <table border="1" cellpadding="5" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap rowspan="2">Account</th>
        <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap colspan="2">Course </th>
        <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap rowspan="2"># Learners<br>Completed</th>
      </tr>
      <tr>
        <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap>Code</th>
        <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left" nowrap>Title</th>
      </tr>
      <%
          Dim vCustIdPrev 
          sOpenDb
          vSql = "SELECT * FROM Snap WHERE UserNo = " & svMembNo & " ORDER BY ParentId, CustId "
          Set oRs = oDb.Execute(vSql)
          Do While Not oRs.Eof
      %>
      <tr>
        <td align="center">
        <% 
           If vCustIdPrev <> oRs("CustId") Then 
        %> 
             <%= fIf(oRs("CustId") = svCustId, "<b>", "")%> <%=oRs("CustId")%> <%= fIf(oRs("CustId") = svCustId, "</b>", "")%> 
        <% 
             vCustIdPrev = oRs("CustId")
           Else 
        %>
            &nbsp;&nbsp;&nbsp; 
        <% 
           End If
        %> 
        </td>
        <td align="center"><%=oRs("ProgId")%></td>
        <td align="left"><%=oRs("Title")%></td>
        <td align="center"><%=oRs("Completed")%></td>
      </tr>
      <% 
            oRs.MoveNext
          Loop
          Set oRs = Nothing
          sCloseDb  
      %>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
</body>
</html>


