<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<html>

<head>
  <title>EcomReport1</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <base target="_self">
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Launch.js"></script>
  <style>
    .table.main tr:nth-child(odd) { background-color: #eee; }
    .table.main th,
    .table.main td { text-align: left; width: 30px; }
  </style>
</head>

<body>

  <% 
    Server.Execute vShellHi 

    Dim vStrDate, vEndDate, vAddr, vName, vButton, vZero
    vButton = "Online Report"
    vZero   = fDefault(Request("vZero"), "0")
      
   '...if Excel then go to Excel version
    If Request.Form("bExcel").Count = 1 Then 
      Response.Redirect "EcomReport1X.asp?vStrDate=" & Server.UrlEncode(Request.Form("vStrDate")) & "&vEndDate=" & Server.UrlEncode(Request.Form("vEndDate")) & "&vZero=" & vZero
    End If
    
    '...If first pass then display the drop down form and prepopulate the dates
    If Request.Form("vHidden").Count = 0 Then
      vStrDate  = Request("vStrDate")      : If Len(vStrDate) = 0 Then vStrDate = fFormatSqlDate(MonthName(Month(Now)) & " 1, " & Year(Now))
      vEndDate  = Request("vEndDate")      : If Len(vEndDate) = 0 Then vEndDate = fFormatSqlDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))

  %>

  <h1>Ecommerce Sales Summary</h1>
  <h2>This report produces a list of ecommerce buyers from <b><%=Left(svCustId, 4) %></b> sites with their purchases and address which can be output online or in Excel.</h2>

  <form method="POST" action="EcomReport1.asp">
    <input type="Hidden" name="vHidden" value="Hidden">
    <table style="margin-top: 30px;" class="table">
      <tr>
        <th>Select Start Date :</th>
        <td>
          <input type="text" name="vStrDate" size="11" value="<%=vStrDate%>">
          <a title="Empty field" class="debug" onclick="emptyField('vStrDate')" href="#">&#937;</a>&nbsp; ie Jan 1, 2005 (mmm d, yyyy). </td>
      </tr>
      <tr>
        <th>Select End Date :</th>
        <td>
          <input type="text" name="vEndDate" size="11" value="<%=vEndDate%>">
          <a title="Empty field" class="debug" onclick="emptyField('vEndDate')" href="#">&#937;</a>&nbsp; ie Mar 31, 2005 (mmm d, yyyy).</td>
      </tr>
      <tr>
        <th>Show Zero Dollar Sales :</th>
        <td>
          <input name="vZero" type="radio" value="1" <%=fcheck("1", vZero)%> />Yes 
            <input name="vZero" type="radio" value="0" <%=fcheck("0", vZero)%> />No
        </td>
      </tr>
      <tr>
        <th>Then click either :</th>
        <td>
          <input type="submit" value="Online" name="bOnline" id="bOnline1" class="button">&nbsp;&nbsp; or&nbsp;
          <input type="submit" value="Excel" name="bExcel" class="button"></td>
      </tr>
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
    </table>
  </form>

  <%
    Else
  %>


  <h1>Ecommerce Sales Summary</h1>
  <h2>These details reflect the price paid for each program.&nbsp; They do not include taxes or factor in royalties.<br>
    Parent Account is the Customer ID of the Site making the sale.&nbsp; New Account is the ID that was created for a Group Site.<br>
    Source is either &quot;E&quot; (normal ecommerce), &quot;C&quot; (manual payments to customer) or &quot;V&quot; (manual payments to Vubiz)<br>
  </h2>


  <table style="margin-top: 30px;" class="table main">
    <tr>
      <th class="rowshade">Parent Account</th>
      <th class="rowshade">New Account</th>
      <th class="rowshade" style="white-space: nowrap;">Order/Expires</th>
      <th class="rowshade">Order Id</th>
      <th class="rowshade">Source</th>
      <th class="rowshade">Name</th>
      <th class="rowshade">Organization</th>
      <th class="rowshade">Mailing Address</th>
      <th class="rowshade">Programs</th>
      <th class="rowshade">Quantity</th>
      <th class="rowshade">Price</th>
    </tr>
    <%
      vStrDate = fFormatDate(fDefault(Request("vStrDate"), "Jan 1, 2000"))
      vEndDate = fFormatDate(fDefault(Request("vEndDate"), Now))

      vSql = " SELECT * FROM Ecom WHERE " _
           & "   Ecom_Issued >= '" & vStrDate & "' AND Ecom_Issued < DATEADD(d, 1, '" & vEndDate & "')" _
           & "   AND (LEFT(Ecom_CustId, 4) = '" & Left(svCustId, 4) & "')" _
           &     fIf (vZero = 0, "   AND Ecom_Amount <> 0", "") _
           & " ORDER BY Ecom_CustId, Ecom_Issued, Ecom_LastName, Ecom_FirstName "
'     sDebug

      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRS.eof
        sReadEcom
        vName = vEcom_FirstName & " " & vEcom_LastName
        vAddr = vEcom_Address & "<br>" & vEcom_City & ", " & vEcom_Postal & ", " & vEcom_Province & ", " & vEcom_Country & "<br>" & vEcom_Phone  & "<br>" & vEcom_Email 
    %>
    <tr>
      <td><%=vEcom_CustId%></td>
      <td><%=fIf(Len(Trim(vEcom_NewAcctid)) > 0 , Left(vEcom_CustId, 4), "") & vEcom_NewAcctId%></td>
      <td style="white-space: nowrap;"><%=fFormatDate(vEcom_Issued)%>
        <br />
        <%=fFormatDate(vEcom_Expires)%></td>
      <td><%=vEcom_OrderId%></td>
      <td style="text-align: center"><%=vEcom_Source%></td>
      <td><%=fLeft(vName, 24)%></td>
      <td><%=vEcom_Organization%></td>
      <td><%=vAddr%></td>
      <td><%=vEcom_Programs%></td>
      <td><%=vEcom_Quantity%></td>
      <td><%=FormatCurrency(vEcom_Prices, 2)%></td>
    </tr>
    <%
        oRs.MoveNext	        
      Loop
      sCloseDB
    %>
    <tr>
      <td colspan="13" style="text-align: center">&nbsp;<p>
        <input onclick="location.href = 'EcomReport1.asp'" type="button" value="Return" name="bReturn" id="bReturn" class="button085"><%=f10%>
        <!--        <input type="button" onclick="jPrint();" value="Print" name="bOnline" id="bOnline2" class="button085">-->
      </p>
        <p>&nbsp;</p>
      </td>
    </tr>
  </table>
  <%
    End If
  %>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


