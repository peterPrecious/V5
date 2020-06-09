<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi
    
    Dim vAcct, vSort

    vAcct    = fDefault(Request("vAcct"), "c")
    vSort    = fDefault(Request("vSort"), "u")

    If Request.Form("bExcel").Count = 1 Then   
      Response.Redirect "LogReport3_X.asp?vAcct=" & Request.Form("vAcct") & "&vSort=" & Request.Form("vSort")
    ElseIf Request.Form("bOnline").Count = 0 Then   
  %>


  <form method="POST" action="LogReport3.asp">
    <input type="hidden" name="vHidden" value="Hidden">
    <table border="1" width="100%" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <td colspan="2" align="center">
        <h1 align="center">Module Usage Report</h1>
        <h2>This report displays the Time Spent (in minutes) for each of your modules</h2>
        </td>
      </tr>
      <% If svMembLevel = 5 Then %>
      <tr>
        <th align="right" nowrap width="35%" valign="top">Select for Accounts :</th>
        <td width="65%" valign="top">
          <input type="radio" value="*" name="vAcct">All (Administrator only)<br> 
          <input type="radio" value="c" checked name="vAcct">Current Account
        </td>
      </tr>
      <% End If %>
      <tr>
        <th align="right" nowrap width="35%" valign="top">Sort Report by :</th>
        <td width="65%" valign="top">
          <input type="radio" name="vSort" value="u" <%=fCheck("u", vSort)%>>Usage (most popular)<br> 
          <input type="radio" name="vSort" value="t" <%=fCheck("t", vSort)%>>Module Title<br> 
          <input type="radio" name="vSort" value="i" <%=fCheck("i", vSort)%>>Module Id</td>
      </tr>
      <tr>
        <th height="100" width="100%" colspan="2">Select one of...<br><br>
          <input type="submit" value="Online" name="bOnline" id="bOnline" class="button070"><%=f10%>
          <input type="submit" value="Excel" name="bExcel" class="button070">
        </th>
      </tr>
    </table>
  </form>


<% Else %>

  <div align="center">
    <table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <th bordercolor="#FFFFFF" colspan="3">
          <h1>Module Access Report - Count</h1>
          <h2 align="left">This report displays the Time Spent (in minutes) for each of your modules sorted by <%=fIf(vSort="u","Time Spent", fIf(vSort="t", "Title", "Module Id"))%>.</h2>
        </th>
      </tr>
      <tr>
        <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Module Id</th>
        <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Title</th>
        <th align="right" nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF">Time Spent<br>(Minutes)&nbsp; </th>
      </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <%
          vSql = " SELECT Mods_Id, Mods_Title, SUM(CAST(RIGHT(Logs_Item, 6) AS int)) AS [TimeSpent]"_
               & " FROM Logs INNER JOIN V5_Base.dbo.Mods ON SUBSTRING(Logs.Logs_Item, 9, 6) = V5_Base.dbo.Mods.Mods_ID"_
               & " WHERE (Logs_Type = 'P')"_
               &     fIf(vAcct = "c",       " AND (Logs_AcctId = '"  & svCustAcctId & "')", "")_
               & " GROUP BY Mods_Id, Mods_Title"
          If vSort = "u" Then
            vSql = vSql & " ORDER BY TimeSpent DESC, Mods_Title, Mods_Id"
          ElseIf vSort = "t" Then
            vSql = vSql & " ORDER BY Mods_Title, Mods_Id"
          Else
            vSql = vSql & " ORDER BY Mods_Id"
          End If

'         sDebug
      
          sOpenDb
          Set oRs = oDb.Execute(vSql)
          Do While Not oRS.Eof
            If oRs("TimeSpent") > 0 Then
      %>
      <tr>
        <td><%=oRs("Mods_Id")%>&nbsp; </td>
        <td><%=oRs("Mods_Title")%></td>
        <td align="right"><%=oRs("TimeSpent")%>&nbsp; </td>
      </tr>
      <%
          End If
          oRs.MoveNext	        
        Loop
        sCloseDB
	    %>
      <tr>
        <td colspan="3" align="center" bordercolor="#DDEEF9" height="60">&nbsp;
          <a href="LogReport3.asp?vAcct=<%=vAcct%>&vSort=<%=vSort%>">
          <img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif">
          </a>
        </td>
      </tr>
    </table>

  <%
      End If 
    Server.Execute vShellLo 
  %>

</body>

</html>