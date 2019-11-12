<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <title>ChannelExpiryReport</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
  	Server.Execute vShellHi	

    If Request("bExcel").Count > 0 Then Response.Redirect "CustomerExpiryReport_X.asp"
  %>

  <h1>Channel Expiry Report</h1>
  <h2>This Report lists active Child Accounts of Parent <%=svCustId %> that will expire within the next twelve months.</h2>
  <h3>The default (Online) option lists a maximum of 500 accounts, Excel will list up to 5000 accounts.</h3>
  <div style="text-align: center; margin:30px;">
    <input type="submit" value="Online" name="bOnline" class="button">
    <input type="button" value="Excel" name="bExcel" class="button" onclick="location.href= 'CustomerExpiryReport_X.asp'">
  </div>
  <table>
    <tr>
       <th class="rowshade" style="width:10%; text-align: left;">Customer Id</th>
       <th class="rowshade" style="width:70%; text-align: left;">Title</th>
       <th class="rowshade" style="width:10%; text-align: center;"># Learners</th>
       <th class="rowshade" style="width:10%; text-align: center;">Expiry Date</th>
       <th class="rowshade" style="width:10%; text-align: center;">Facilitator Id</th>
       <th class="rowshade" style="width:10%; text-align: center;">Facilitator Email</th>
       <th class="rowshade" style="width:10%; text-align: center;">Facilitator Last Visit</th>
    </tr>
    <%

      vSql = " " _
           & "SELECT TOP 500 "_
           & "	Cust_Id AS CustId, "_ 
           & "  Cust_Title AS CustTitle, "_ 
           & "	Count(Memb_Id) AS MembCount, "_ 
           & "  Cust_Expires AS CustExpires "_
           & "FROM "_ 
           & "  Cust INNER JOIN "_ 
           & "	Memb ON Cust_AcctId = Memb_AcctId "_ 
           & "WHERE "_ 
           & "	(Cust_ParentId = '" & RIGHT(svCustId, 4) & "') AND "_ 
           & "	(Cust_Active = 1) AND "_ 
           & "	(Memb_Level < 4) AND "_   
           & "	(Cust_Expires > getDate()) AND "_
           & "	(Cust_Expires < DATEADD(year, 1, getDate())) "_
           & "GROUP BY "_ 
           & "  Cust_Id, "_
           & "  Cust_Title, "_ 
           & "  Cust_Expires "_
           & "ORDER BY  "_
           & "  Cust_Expires "

      vSql = " " _
           & "SELECT TOP 500 "_
	         & "  cu.Cust_Id AS custId, "_
	         & "  cu.Cust_Title AS custTitle, "_
	         & "  Count(m1.Memb_Id) AS membCount, "_
	         & "  Cust_Expires AS custExpires, "_
	         & "  m2.Memb_Id AS facId, "_
           & "	m2.Memb_Email AS facEmail, "_
           & "	m2.Memb_LastVisit AS facLast "_
           & "FROM "_  
           & "	V5_Vubz.dbo.Cust cu											                  INNER JOIN "_
           & "	V5_Vubz.dbo.Memb m1 ON cu.Cust_AcctId = m1.Memb_AcctId		INNER JOIN "_
           & "	V5_Vubz.dbo.Memb m2 ON cu.Cust_AcctId = m2.Memb_AcctId "_
           & "WHERE "_  
           & "	(Cust_ParentId = '" & RIGHT(svCustId, 4) & "') AND "_ 
           & "	(cu.Cust_Active = 1) AND "_  
           & "	(m1.Memb_Level < 4) AND "_    
           & "	(m1.Memb_Internal = 0) AND "_
           & "	(cu.Cust_Expires > getDate()) AND "_ 
           & "	(cu.Cust_Expires < DATEADD(year, 1, getDate())) AND "_
           & "	(m2.Memb_Level = 3) AND "_
           & "	(m2.Memb_Internal = 0) AND "_
           & "	(m2.Memb_LastVisit is not null) "_ 
           & "GROUP BY "_  
           & "	cu.Cust_Id, "_ 
           & "	cu.Cust_Title, "_  
           & "	cu.Cust_Expires, "_
           & "	m2.Memb_Id, "_
           & "	m2.Memb_Email, "_
           & "	m2.Memb_LastVisit "_
           & "ORDER BY "_  
           & "	cu.Cust_Expires " 

      sOpenDb      
      Set oRs = oDb.Execute(vSql)      
      Do While Not oRs.Eof

    %>
    <tr>
      <td><%=oRs("custId")%></td>
      <td><%=Left(oRs("custTitle"), 30)%></td>
      <td style="text-align: center;"><%=oRs("membCount")%></td>
      <td style="text-align: center;"><%=fFormatDate(oRs("custExpires"))%></td>
      <td><%=oRs("facId")%></td>
      <td><%=oRs("facEmail")%></td>
      <td style="text-align: center;"><%=fFormatDate(oRs("facLast"))%></td>
    </tr>
    <%  
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDB    
    %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
</body>

</html>
