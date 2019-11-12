<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vStrDate, vEndDate, vCustIdPrev

  vStrDate = "Jan 1, 2000"
  vEndDate = fFormatSqlDate(Now)

  '...create the snapshot file
  sOpenCmd
  With oCmd
    .CommandText = "spSnapshot"
    .Parameters.Append .CreateParameter("@UserNo",    adInteger,  adParamInput,      , svMembNo)
    .Parameters.Append .CreateParameter("@CustId",    adChar,     adParamInput,   008, svCustId)
    .Parameters.Append .CreateParameter("@AcctId",  	adChar,     adParamInput,   004, svCustAcctId)
    .Parameters.Append .CreateParameter("@StrDate",   adDBDate,   adParamInput,      , vStrDate)
    .Parameters.Append .CreateParameter("@EndDate",		adDBDate,   adParamInput,      , vEndDate)
  End With
  oCmd.Execute()
  Set oCmd = Nothing
  sCloseDb
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Calendar.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td valign="top">
      <h1 align="center">Course Usage Snapshot</h1>
      <h2 align="left">This report shows the number of active and inactive Learner Profiles in this <b>Parent</b> account and all <b>Child</b> accounts (the number does NOT include facilitators). It also shows the count of all those who have <b>Started</b> the course (based on whether they have spent time in the course and/or the exam, as well as the number who have <b>Completed</b> the course.</h2>
      </td>
    </tr>
    <tr>
      <td valign="top"><input type="hidden" value="<%=Request("vParmNo")%>" name="vParmNo">
      <div align="center">
        <table border="1" cellspacing="0" cellpadding="5" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap rowspan="2">Account</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap rowspan="2"># Learners<br>Enrolled</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap colspan="2">Course </th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap colspan="2"># Learners</th>
          </tr>
          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap>Code</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left" nowrap>Title</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap>Started</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap>Completed</th>
          </tr>
          <%
              sOpenDb
              vSql = "SELECT * FROM Snap WHERE UserNo = " & svMembNo & " ORDER BY ParentId, CustId "
              Set oRs = oDb.Execute(vSql)
              Do While Not oRs.Eof
          %>
          <tr>

          <% 
              If vCustIdPrev <> oRs("CustId") Then 
          %>        
            <td align="center"><%= fIf(oRs("CustId") = svCustId, "<b>", "")%> <%=oRs("CustId")%> <%= fIf(oRs("CustId") = svCustId, "</b>", "")%> </td>
            <td align="center"><%=fMembCount (oRs("AcctId"))%></td>
          <% 
             vCustIdPrev = oRs("CustId")
             Else 
          %>
            <td align="center" colspan="2">&nbsp;</td>
                   
          <% End If%>        



            <td align="center"><%=oRs("ProgId")%></td>
            <td align="left"><%=oRs("Title")%></td>
            <td align="center"><%=oRs("Started")%></td>
            <td align="center"><%=oRs("Completed")%></td>
          </tr>
          <% 
                oRs.MoveNext
              Loop
              Set oRs = Nothing
              sCloseDb  
          %>
        </table>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

