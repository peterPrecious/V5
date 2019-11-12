<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% 
    Server.Execute vShellHi 
  %>
  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3" cellspacing="0">
    <tr>
      <td colspan="5" align="center"><h1><br>Customer Sellers and Owners Report</h1>
      </td>
    </tr>
    <tr>
      <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF">Cust Id</th>
      <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Web Site</th>
      <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Name</th>
      <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF">Seller</th>
      <th nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF">Owner</th>
    </tr>
    <%
      vSql =        " "
      vSql = vSql & " SELECT DISTINCT Left(Cust_Id, 4) AS [Cust], Cust_EcomSeller, Cust_EcomOwner"
      vSql = vSql & " FROM Cust"
      vSql = vSql & " WHERE Cust_Active = 1 AND (Cust_EcomSeller = 1 OR Cust_EcomOwner = 1) "
      vSql = vSql & " ORDER BY Cust " 
'     sDebug "", vSql

      sOpenDb2
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRS.eof
 
        '...get customer title      
        vSql =        " "
        vSql = vSql & " SELECT TOP 1 Cust_Title, Cust_Url"
        vSql = vSql & " FROM Cust"
        vSql = vSql & " WHERE Cust_Active = 1 AND (Cust_EcomSeller = 1 OR Cust_EcomOwner = 1) AND Left(Cust_Id, 4) = '" & oRs("Cust") & "'"
'       sDebug "", vSql
        Set oRs2 = oDb2.Execute(vSql)

    %> 
    
    <tr>
      <td valign="top" align="center"><%=oRs("Cust")%></td>
      <td valign="top" align="left">
      <% If Len(Trim(oRs2("Cust_Url"))) > 0 Then %>
      <a target="_blank" href="<%=oRs2("Cust_Url")%>"><%=oRs2("Cust_Url")%></a>
      <% End If %> 
      </td>
      <td valign="top" align="left"><%=oRs2("Cust_Title")%></td>
      <td valign="top" align="center"><%=fIf(oRs("Cust_EcomSeller") = True, "<img border='0' src='../Images/Common/CheckMark.jpg' width='12' height='15'>", "")%></td>
      <td valign="top" align="center"><%=fIf(oRs("Cust_EcomOwner") = True, "<img border='0' src='../Images/Common/CheckMark.jpg' width='12' height='15'>", "")%></td>
    </tr>
    <%
        Set oRs2 = Nothing
        oRs.MoveNext	        
      Loop
      sCloseDB
    %> 
    
    <tr>
      <td colspan="5" align="center">
        &nbsp;<p>
        <input type="submit" onclick="window.open('CustomerSellerOwners_x.asp')" value="Excel" name="bExcel1" class="button"></p><p>
        </p>
      </td>
    </tr>
  </table>
  <%
    Server.Execute vShellLo 
  %>

</body>

</html>