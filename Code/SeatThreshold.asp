<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim status : status = 0

  '... every morning a small job runs to create the current "apps.dbo.threshold" table

  If Request("children").Count = 1 Then
    status = 2
    Dim acctId, title, rows
    acctId = Split(Request("children"),"|")(0)
    title = Split(Request("children"),"|")(1)

    vSql = " SELECT  [custId] "_
         & "        ,[expires] "_
         & "        ,[purchased] "_
         & "        ,[assigned] "_
         & "        ,[usage] "_
         & "        ,[facilitator] "_
         & " FROM [apps].[dbo].[threshold] "_
         & " WHERE parent = '" & RIGHT(acctId, 4) & "' and usage > 75 "
    rows = ""
  
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.EOF 
      rows = rows & "   <tr>" & vbCrLf
      rows = rows & "     <td><a target='_blank' href='ProgramsAssigned.asp?custId=" & oRs("custId") & "'>" & oRs("custId") & "</a></td>" & vbCrLf
      rows = rows & "     <td>" & oRs("expires") & "</td>" & vbCrLf
      rows = rows & "     <td>" & oRs("purchased") & "</td>" & vbCrLf
      rows = rows & "     <td>" & oRs("assigned") & "</td>" & vbCrLf
      rows = rows & "     <td>" & fFormatNumber (oRs("usage"), 1) & "</td>" & vbCrLf
      rows = rows & "     <td>" & oRs("facilitator") & "</td>" & vbCrLf
      rows = rows & "   </tr>" & vbCrLf 
    oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb

  Else
    status = 1
    Dim options
 
    '...build the drop down to select parent
    vSql = " SELECT DISTINCT "_
         & "   c2.Cust_Id     AS custId, "_
         & "   c2.Cust_Title  AS custTitle "_
         & "FROM "_ 
         & "   [V5_Vubz].[dbo].[Cust] c1 INNER JOIN "_ 
         & "   [V5_Vubz].[dbo].[Cust] c2 ON c2.Cust_Id = LEFT(c1.Cust_Id, 4) + c1.Cust_ParentId "_ 
         & "WHERE "_ 
         & "   LEN(c1.Cust_ParentId) = 4 AND "_ 
         & "   c1.Cust_Expires > getdate() AND "_
         & "   c1.Cust_ParentId IN (SELECT parent FROM (SELECT parent FROM [apps].[dbo].[threshold] WHERE purchased > 1 AND usage >= 75) a GROUP BY parent) "_
         & "ORDER BY "_
         & "   c2.Cust_Id "

    options = "<select id='children' name='children' size='10' style='width:90%'>" & vbCrLf

    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.EOF 
      options = options & "      <option value='" & oRs("custId") + "|" & oRs("custTitle") & "'>" & oRs("custId") & " - " & oRs("custTitle") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb

    options = options & "    </select>"

  End If



%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <title></title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <style>
    #children tr th, #children tr td {
      text-align: center;
    }

    #children tr th {
      width: 19%;
    }
    #children tr th:nth-child(6), #children tr td:nth-child(6)  {
      text-align:left;
    }
  </style>
</head>

<body>

  <% 
    Server.Execute vShellHi
  %>

  <h1>This lists Child Accounts where at least 75% of seats have been assigned.</h1>


  <% If status = 1 Then %>

  <h2>Select the Parent Account</h2>

  <form action="SeatThreshold.asp" style="width: 600px; margin: 30px auto;">
    <%=options %>
    <button class="button" type="submit">GO</button>
  </form>

  <% Else  %>

  <h2>Parent : <%=acctId & " - " & title %></h2>
  <h3>Click on a Child Account for more Details</h3>

  <table id="children" class="table" style="width: 700px; margin: 50px auto;">
    <tr>
      <th class="rowshade">Account</th>
      <th class="rowshade">Expires</th>
      <th class="rowshade">Purchased</th>
      <th class="rowshade">Assigned</th>
      <th class="rowshade">Usage %</th>
      <th class="rowshade">Facilitator</th>
    </tr>
    <%=rows%>
  </table>

  <div style="text-align: center; margin: 50px;">
    <button onclick="history.back(1)" class="button" type="submit">Back</button>
  </div>

  <% End If %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
</body>

</html>


