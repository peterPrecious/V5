<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<% 
	Response.Buffer = true

  Dim vLastName, vDays, vAction, vEcomNos, aEcomNos
  vLastName  = fDefault(Request("vLastName"), "")
  
  vAction  = fDefault(Request("vAction"), 0)
  vDays    = fDefault(Request("vDays"),0)
  vEcomNos = Replace(fDefault(Request("vEcomNos"), ""), " ", "")

  If Request("bUpdate").Count > 0 Then
    If vAction <> 0 And vDays <> 0 And vEcomNos <> "" Then
      aEcomNos = Split(vEcomNos, ",")
      sOpenDb
      For i = 0 To Ubound(aEcomNos)
        vSql = "UPDATE Ecom SET Ecom_Expires = DATEADD([day], " & vDays * vAction & ", Ecom_Expires) WHERE (Ecom_No = " & aEcomNos(i) & ")"
  '     Response.Write "<br>" & vSql
        oDb.Execute(vsql)
      Next    
      sCloseDb
    End If
  End If
%>

<html>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <script src="/V5/Inc/jQuery.js"></script>
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <title>Ecommerce Report</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <form method="POST" action="EcomExtend.asp">
    <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
      <input type="Hidden" name="vHidden" value="Hidden"><tr>
        <td><h1 align="center">Ecommerce Extension Report</h1><h2>This enables you to extend (or reduce) access for an ecommerce transaction.&nbsp; Enter all or part of the cardholder name / learner&#39;s last name then click <b>GO</b>.&nbsp;&nbsp; When you have identified the learner then select the Program(s) and the number of days of extension (or reduction) then click <b>Update</b>. <font color="#FF0000">&nbsp;Note that only the first 100 names are listed based on your selection.</font></h2></td>
      </tr>
      <tr>
        <th valign="top">Cardholder Name /Learner Surname (ie Smith, Sm) :&nbsp; <input type="text" name="vLastName" size="15" value="<%=vLastName%>"> <input type="submit" value="Go" name="bGo" class="button"><p>&nbsp;</p>
        </th>
      </tr>
    </table>
    <table border="1" cellpadding="2" cellspacing="0" bordercolor="#DDEEF9" id="table1" width="100%" style="border-collapse: collapse">
      <tr>
        <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Cardholder</th>
        <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">First Name</th>
        <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Last Name</th>
        <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Program</th>
        <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Title</th>
        <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">$Amount</th>
        <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Ordered</th>
        <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Expires</th>
        <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Extend?</th>
      </tr>
      <% 
        Dim vTitle
        vSql = "SELECT TOP 100 Ecom.Ecom_No, Ecom.Ecom_CardName, Ecom.Ecom_FirstName, Ecom.Ecom_LastName, Ecom.Ecom_Programs, Ecom.Ecom_Issued, Ecom.Ecom_Prices, Ecom.Ecom_Expires, V5_Base.dbo.Prog.Prog_Title1 " _
             & "FROM Ecom INNER JOIN V5_Base.dbo.Prog ON Ecom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id " _
             & "WHERE (Ecom.Ecom_AcctId = '" & svCustAcctId & "') AND ((Ecom.Ecom_LastName LIKE '%" & vLastName & "%') OR (Ecom.Ecom_CardName LIKE '%" & vLastName & "%')) " _
             & "ORDER BY Ecom.Ecom_LastName, Ecom.Ecom_Issued DESC"
'       sDebug     
        sOpenDb
        Set oRs = oDb.Execute(vsql)
        Do While Not oRs.Eof
          vTitle = oRs("Prog_Title1")
          If Instr(vTitle, "<") > 0 Then vTitle = Left(vTitle, Instr(vTitle, "<") - 1)
      %> 
      <tr>
        <td><%=oRs("Ecom_Cardname")%></td>
        <td><%=oRs("Ecom_FirstName")%></td>
        <td><%=oRs("Ecom_LastName")%></td>
        <td><%=oRs("Ecom_Programs")%></td>
        <td><%=fLeft(vTitle, 32)%></td>
        <td align="center"><%=FormatCurrency(oRs("Ecom_Prices"),2)%></td>
        <td align="center" nowrap><%=fFormatDate(oRs("Ecom_Issued"))%></td>
        <td align="center" nowrap><%=fFormatDate(oRs("Ecom_Expires"))%></td>
        <td align="center" bgcolor="#DDEEF9" bordercolor="#FFFFFF">
          <input type="checkbox" name="vEcomNos" value="<%=oRs("Ecom_No")%>"></td>
      </tr>
      <%
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
     %>
    </table>
    <div align="center"><h2 class="c6">Make a written note of the Expiry Date before you make any modifications, in case you need to reset your actions.<br>Once you click <b>Update</b> the list will re-appear with the modified dates.</h2>
      <table border="1" cellpadding="5" cellspacing="0" bordercolor="#DDEEF9" id="table2" style="border-collapse: collapse">
        <tr>
          <th>
          <select size="1" name="vAction">
            <option <%=fSelect("1", vAction)%> value="1">Extend</option>
            <option <%=fSelect("-1", vAction)%> value="-1">Reduce</option>
          </select> access to the selected programs by
          <select size="1" name="vDays">
            <option value="0">Select</option>
            <option <%=fSelect("1", vDays)%>>1</option>
            <option <%=fSelect("2", vDays)%>>2</option>
            <option <%=fSelect("3", vDays)%>>3</option>
            <option <%=fSelect("4", vDays)%>>4</option>
            <option <%=fSelect("5", vDays)%>>5</option>
            <option <%=fSelect("6", vDays)%>>6</option>
            <option <%=fSelect("7", vDays)%>>7</option>
            <option <%=fSelect("14", vDays)%>>14</option>
            <option <%=fSelect("21", vDays)%>>21</option>
            <option <%=fSelect("30", vDays)%>>30</option>
            <option <%=fSelect("60", vDays)%>>60</option>
            <option <%=fSelect("90", vDays)%>>90</option>
            <option <%=fSelect("180", vDays)%>>180</option>
            <option <%=fSelect("365", vDays)%>>365</option>
          </select> days
          <input type="submit" value="Update" name="bUpdate" class="button"></th>
        </tr>
      </table>
      <h2 class="c6">&nbsp;</h2>
    </div>
    <p>&nbsp;</p>
  </form>
  <% Server.Execute vShellLo %>

</body>

</html>

