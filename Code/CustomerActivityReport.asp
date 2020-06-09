<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<% 
  Dim vSort, vLearners, vInActive
  vSort     = fDefault(Request("vSort"), "a")  
  vLearners = fDefault(Request("vLearners"), "n")  
  vInactive = fDefault(Request("vInactive"), "n")  
%>

<html>

<head>
  <title>CustomerActivityReport</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <style>
    td { text-align: right; padding-right: 18px; }
      td:nth-child(1) { text-align: center; }
      th:nth-child(1) { text-align: center; }
    .rowshade { text-align: right; }
      .rowshade a { color: white; }
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Customer Activity Report</h1>
  <h2>Active Learners are those have visited the site within the last 12 months, the balance being Inactive Learners. Click on any title to sort the report (sorted column title will appear green).</h2>
  <h3>Include Inactive Accounts? <b><a href="CustomerActivityReport.asp?vSort=<%=vSort%>&vLearners=<%=vLearners%>&vInactive=y"><font color="<%=fIf(vInactive="y","#008000", "#000080")%>">Yes</font></a>| <a href="CustomerActivityReport.asp?vSort=<%=vSort%>&vInactive=n"><font color="<%=fIf(vInactive="n","#008000", "#000080")%>">No</font></a></b>&nbsp;&nbsp;&nbsp;&nbsp;Include Only Learners? <b><a href="CustomerActivityReport.asp?vSort=<%=vSort%>&vLearners=y&vInactive=<%=vInactive%>"><font color="<%=fIf(vLearners="y","#008000", "#000080")%>">Yes</font></a>| <a href="CustomerActivityReport.asp?vSort=<%=vSort%>&vLearners=n&vInactive=<%=vInactive%>"><font color="<%=fIf(vLearners="n","#008000", "#000080")%>">No</font></a></b></h3>


  <table style="width:80%; margin: auto;">
    <tr>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=c&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">Customer</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=a&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">Active <br>Learners</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=i&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">Inactive <br>Learners</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=p&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">% Active <br>Learners</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=t&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">Total <br>Learners</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=v&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>"># Active <br>Visits</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=y&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">Visits<br>/Learner</a></th>
      <th class="rowshade"><a href="CustomerActivityReport.asp?vSort=s&vInactive=<%=vInactive%>&vLearners=<%=vLearners%>">% Server<br>Usage</a></th>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <%
      '...get the customers record set
      Dim vCntActive, vCntInactive, vCntVisits, vCntTotal, vPerActive, vPerVisits, vPerServer
      Dim vCntActiveTot, vCntInactiveTot, vCntTotalTot, vCntVisitsTot, vCntServerTot

      sOpenDb

      '...first get the total number of visits in order to provide the % Server Usager
      vSql = " SELECT "

'     vSql = vSql & " SUM(Memb.Memb_NoVisits) AS Total_Visits "

      vSql = vSql & " SUM(CASE WHEN Memb_Active = 1 AND (Memb_LastVisit >= '" & DateAdd("yyyy", -1, Now) & "' AND Memb_LastVisit IS NOT NULL) THEN Memb_NoVisits ELSE 0 END) AS Total_Visits" 

      vSql = vSql & " FROM Memb WITH (nolock) INNER JOIN Cust ON RIGHT(Memb.Memb_AcctId, 4) = RIGHT(Cust.Cust_Id, 4)"
      If vLearners = "y" Then
        vSql = vSql & " WHERE Memb_Level = 2"
      End If
      If vInactive = "n" Then
        vSql = vSql & " AND Cust_Active = 1"
      End If 

'     sDebug
      Set oRs = oDb.Execute(vSql)
      vCntVisitsTot   = vCntVisitsTot   + oRs("Total_Visits")

      '...now get the details for each account
      vSql = " "
      vSql = vSql & " SELECT LEFT(Cust.Cust_Id, 4) AS Customer, COUNT(Memb_No) AS Total_Users,"
      vSql = vSql & " SUM(CASE WHEN Memb_Active = 1 AND (Memb_LastVisit >= '" & DateAdd("yyyy", -1, Now) & "' AND Memb_LastVisit IS NOT NULL) THEN 1 ELSE 0 END) AS Total_Active," 
      vSql = vSql & " SUM(CASE WHEN Memb_Active = 0 OR Memb_LastVisit IS NULL OR Memb_LastVisit <= '" & DateAdd("yyyy", -1, Now) & "' THEN 1 ELSE 0 END) AS Total_InActive,"
      vSql = vSql & " SUM(CASE WHEN Memb_Active = 1 AND (Memb_LastVisit >= '" & DateAdd("yyyy", -1, Now) & "' AND Memb_LastVisit IS NOT NULL) THEN Memb_NoVisits ELSE 0 END) AS Total_Visits," 
      vSql = vSql & " CAST(CAST(SUM(CASE WHEN Memb_Active = 1 AND (Memb_LastVisit >= '" & DateAdd("yyyy", -1, Now) & "' AND Memb_LastVisit IS NOT NULL) THEN Memb_NoVisits ELSE 0 END) AS FLOAT) / CAST(COUNT(Memb.Memb_No) AS FLOAT) AS FLOAT) AS Percent_Visits, "
      vSql = vSql & " CAST(CAST(SUM(CASE WHEN Memb_Active = 1 AND (Memb_LastVisit >= '" & DateAdd("yyyy", -1, Now) & "' AND Memb_LastVisit IS NOT NULL) THEN 1 ELSE 0 END) AS FLOAT) / CAST(COUNT(Memb_No) AS FLOAT) AS FLOAT) AS Percent_Active,"
      vSql = vSql & " CAST(CAST(SUM(CASE WHEN Memb_Active = 1 AND (Memb_LastVisit >= '4/10/2005 5:48:08 PM' AND Memb_LastVisit IS NOT NULL) THEN Memb_NoVisits ELSE 0 END) AS FLOAT) / CAST(159386 AS FLOAT) AS FLOAT) AS Percent_Server "

      vSql = vSql & " FROM Memb WITH (nolock) INNER JOIN Cust ON RIGHT(Memb.Memb_AcctId, 4) = RIGHT(Cust.Cust_Id, 4)"
      If vLearners = "y" Then
        vSql = vSql & " WHERE Memb_Level = 2"
      End If
      If vInactive = "n" Then
        vSql = vSql & " AND Cust_Active = 1"
      End If        

      vSql = vSql & " GROUP BY LEFT(Cust.Cust_Id, 4)"
      Select Case vSort
        Case "c" : vSql = vSql & " ORDER BY LEFT(Cust.Cust_Id, 4)"
        Case "a" : vSql = vSql & " ORDER BY Total_Active DESC"
        Case "i" : vSql = vSql & " ORDER BY Total_Inactive DESC"
        Case "t" : vSql = vSql & " ORDER BY Total_Users DESC"
        Case "p" : vSql = vSql & " ORDER BY Percent_Active DESC"
        Case "v" : vSql = vSql & " ORDER BY Total_Visits DESC"
        Case "y" : vSql = vSql & " ORDER BY Percent_Visits DESC"
        Case "s" : vSql = vSql & " ORDER BY Percent_Server DESC"
      End Select
      
     sDebug
      Set oRs = oDb.Execute(vSql)

      Do While Not oRs.Eof
        vCntActive      = oRs("Total_Active")
        vCntInactive    = oRs("Total_Inactive")
        vCntVisits      = oRs("Total_Visits")
        vCntTotal       = oRs("Total_Users")
        vPerActive      = oRs("Percent_Active")
        vCntVisits      = oRs("Total_Visits")
        vPerVisits      = oRs("Percent_Visits")
        vPerServer      = oRs("Percent_Server")

        vCntActiveTot   = vCntActiveTot   + vCntActive
        vCntInactiveTot = vCntInactiveTot + vCntInactive
        vCntTotalTot    = vCntTotalTot    + vCntTotal
        vCntServerTot   = vCntServerTot   + vPerServer
    %>
    <tr>
      <td><%=oRs("Customer")%> </td>
      <td><%=vCntActive%> </td>
      <td><%=vCntInactive%> </td>
      <td><%=FormatPercent(vPerActive, 0)%> </td>
      <td><%=vCntTotal%> </td>
      <td><%=vCntVisits%> </td>
      <td><%=FormatNumber(vPerVisits, 0)%> </td>
      <td><%=FormatPercent(vPerServer, 0)%> </td>
    </tr>
    <%
        oRs.MoveNext
      Loop
      sCloseDB
    %>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td class="rowshade"><%=vCntActiveTot%> </td>
      <td class="rowshade"><%=vCntInactiveTot%> </td>
      <td class="rowshade"><%=FormatPercent(vCntActiveTot/vCntTotalTot, 0)%> </td>
      <td class="rowshade"><%=vCntTotalTot%> </td>
      <td class="rowshade"><%=vCntVisitsTot%> </td>
      <td class="rowshade"><%=FormatNumber(vCntVisitsTot/vCntTotalTot, 0)%> </td>
      <td class="rowshade"><%=FormatPercent(vCntServerTot, 0)%> </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


