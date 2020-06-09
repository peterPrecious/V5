<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <base target="_self">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Launch.js"></script>
</head>

<body>

  <% 
    Server.Execute vShellHi 

    Dim vStrDate, vEndDate, vAddr, vName, vButton
    vButton = "Online Report"
      
   '...if Excel then go to Excel version
    If Request.Form("bExcel").Count = 1 Then 
      Response.Redirect "EcomReport1X.asp?vStrDate=" & Server.UrlEncode(Request.Form("vStrDate")) & "&vEndDate=" & Server.UrlEncode(Request.Form("vEndDate"))
    End If
    
    '...If first pass then display the drop down form and prepopulate the dates
    If Request.Form("vHidden").Count = 0 Then
      vStrDate = Left(fFormatSqlDate(Now), 4) & "01" & Mid(fFormatSqlDate(Now), 7)
      vEndDate = Left(fFormatSqlDate(Now), 4) & "31" & Mid(fFormatSqlDate(Now), 7)
      If Not IsDate(vEndDate) Then
        vEndDate = Left(fFormatSqlDate(Now), 4) & "30" & Mid(fFormatSqlDate(Now), 7)
        If Not IsDate(vEndDate) Then
          vEndDate = Left(fFormatSqlDate(Now), 4) & "29" & Mid(fFormatSqlDate(Now), 7)
          If Not IsDate(vEndDate) Then
            vEndDate = Left(fFormatSqlDate(Now), 4) & "28" & Mid(fFormatSqlDate(Now), 7)
          End If
        End If
      End If   
  %>
  <form method="POST" action="EcomReport2.asp">
    <input type="Hidden" name="vHidden" value="Hidden">
    <table border="1" width="100%" cellpadding="5" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td colspan="2" align="center">
        <h1><br>Ecommerce Sales Status</h1>
        <h2>This report produces a list of ecommerce buyers from <b><%=Left(svCustId, 4) %></b> sites with their purchases and address which can be output online or in Excel.<br>&nbsp; </h2>
        </td>
      </tr>
      <tr>
        <th align="right" valign="top" width="30%">Select Start Date :</th>
        <td width="70%"><input type="text" name="vStrDate" size="11" value="<%=vStrDate%>"> <a title="<!--webbot bot='PurpleText' PREVIEW='Empty field'--><%=fPhra(000943)%>" class="debug" onclick="emptyField('vStrDate')" href="#">&#937;</a>&nbsp; ie Jan 1, 2005 (mmm d, yyyy). </td>
      </tr>
      <tr>
        <th align="right" valign="top" width="30%">Select End Date :</th>
        <td width="70%"><input type="text" name="vEndDate" size="11" value="<%=vEndDate%>"> <a title="<!--webbot bot='PurpleText' PREVIEW='Empty field'--><%=fPhra(000943)%>" class="debug" onclick="emptyField('vEndDate')" href="#">&#937;</a>&nbsp; ie Mar 31, 2005 (mmm d, yyyy).</td>
      </tr>
      <tr>
        <th align="right" valign="top" width="30%">Select Program Ids: </th>
        <th align="left" width="70%"><input type="text" name="vPrograms" size="40" value="<%=vPrograms%>"></th>
      </tr>
      <tr>
        <th align="right" valign="top" width="30%">Then click either :</th>
        <th align="left" width="70%"><input type="submit" value="Online" name="bPrint" id="bPrint" class="button"> or ...&nbsp; <input type="submit" value="Excel" name="bExcel" class="button"></th>
      </tr>
      <tr>
        <td colspan="2" align="center">&nbsp;</td>
      </tr>
    </table>
  </form>
  <%
    Else
  %>
  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3" cellspacing="0">
    <tr>
      <td colspan="13" align="center">
      <h1><br>Ecommerce Sales Summary</h1>
      <h2>These details reflect the price paid for each program.&nbsp; They do not include taxes or factor in royalties.<br>Parent Account is the Customer ID of the Site making the sale.&nbsp; New Account is the ID that was created for a Group Site.<br>Source is either &quot;E&quot; (normal ecommerce), &quot;M&quot; (manual payments to customer) or &quot;V&quot; (manual payments to Vubiz)<br></h2>
      </td>
    </tr>
    <tr>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Parent Account</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">New Account</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Order Date</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Expires</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Agent</th>
      <th nowrap align="center" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Source</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Name</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Organization</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Email Address</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Mailing Address</th>
      <th nowrap align="left"   bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Programs</th>
      <th nowrap align="right"  bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Quantity</th>
      <th nowrap align="right"  bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Price</th>
    </tr>
    <%
      vStrDate = fFormatDate(fDefault(Request("vStrDate"), "Jan 1, 2000"))
      vEndDate = fFormatDate(fDefault(Request("vEndDate"), Now))

      vSql = " SELECT *"_
           & " FROM Ecom" _
           & " WHERE Ecom_Issued BETWEEN '" & vStrDate & "' AND '" & vEndDate & "'" _
           & "   AND (LEFT(Ecom_CustId, 4) = '" & Left(svCustId, 4) & "')" _
           & "   AND Ecom_Amount <> 0" _
           & " ORDER BY Ecom_CustId, Ecom_Issued, Ecom_LastName, Ecom_FirstName "

'     sDebug

      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRS.eof
        sReadEcom
        vName = vEcom_FirstName & " " & vEcom_LastName
        vAddr = vEcom_Address & "<br>" & vEcom_City & ", " & vEcom_Postal & ", " & vEcom_Province & ", " & vEcom_Country & "<br>" & vEcom_Phone 
    %>
    <tr>
      <td valign="top"><%=vEcom_CustId%></td>
      <td valign="top"><%=fIf(Len(Trim(vEcom_NewAcctid)) > 0 , Left(vEcom_CustId, 4), "") & vEcom_NewAcctId%></td>
      <td valign="top" nowrap><%=fFormatDate(vEcom_Issued)%></td>
      <td valign="top" nowrap><%=fFormatDate(vEcom_Expires)%></td>
      <td valign="top"><%=vEcom_Agent%></td>
      <td valign="top" align="center"><%=vEcom_Source%></td>
      <td valign="top"><%=fLeft(vName, 24)%></td>
      <td valign="top"><%=vEcom_Organization%></td>
      <td valign="top"><%=vEcom_Email%></td>
      <td valign="top"><%=vAddr%></td>
      <td valign="top"><%=vEcom_Programs%></td>
      <td valign="top" align="right"><%=vEcom_Quantity%></td>
      <td valign="top" align="right"><%=FormatCurrency(vEcom_Prices, 2)%></td>
    </tr>
    <%
        oRs.MoveNext	        
      Loop
      sCloseDB
    %>
    <tr>
      <td colspan="13" align="center">&nbsp;<p><input onclick="history.back(1)" type="button" value="Return" name="bReturn" id="bReturn" class="button085"><%=f10%> <input type="button" onclick="jPrint();" value="<%=bPrint%>" name="bPrint" id="bPrint" class="button085"></p>
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

