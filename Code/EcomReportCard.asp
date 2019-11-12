<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vCredit
  Dim vStrDate, vEndDate, vStrDateErr, vEndDateErr, vUrl

  '...default to previous month
  If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then
    vStrDateErr = "" : vStrDate = fFormatSqlDate(MonthName(Month(Now)) & " 1, " & Year(Now))
    vEndDateErr = "" : vEndDate = fFormatSqlDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))
  Else
    vStrDate  = fFormatSqlDate(Request("vStrDate")) 
    If Request("vStrDate") = "" Then 
      vStrDate = ""
    ElseIf vStrDate = " " Then
      vStrDate  = Request("vStrDate") 
      vStrDateErr = "Error"
    End If
    vEndDate  = fFormatSqlDate(Request("vEndDate"))
    If Request("vEndDate") = "" Then 
      vEndDate = ""
    ElseIf vEndDate = " " Then
      vEndDate  = Request("vEndDate") 
      vEndDateErr = "Error"
    End If
    If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
      If DateDiff("d", vStrDate, vEndDate) < 0 Then
        vEndDateErr = "Error"
      End If
    End If
  End If

  vFind          = fDefault(Request("vFind"), "S")
  vFindId        = fUnQuote(Request("vFindId"))
  vFindFirstName = fUnQuote(Request("vFindFirstName"))
  vFindLastName  = fUnQuote(Request("vFindLastName"))
  vFindEmail     = fNoQuote(Request("vFindEmail"))

  '...processing the form?
  If Request.Form.Count > 0 Then
    '...goto online or excel reports
    vUrl   = "EcomReportCard1.asp" _    
           & "?vStrDate="       & fUrlEncode(vStrDate)       _
           & "&vEndDate="       & fUrlEncode(vEndDate)       _
           & "&vFind="          & fUrlEncode(vFind)          _
           & "&vFindId="        & fUrlEncode(vFindId)        _
           & "&vFindFirstName=" & fUrlEncode(vFindFirstName) _
           & "&vFindLastName="  & fUrlEncode(vFindLastName)  _
           & "&vFindEmail="     & fUrlEncode(vFindEmail)     
'   Response.Write vUrl 
    Response.Redirect vUrl 
  End If
  
%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td valign="top">
      <h1 align="center">Ecommerce Report Card</h1>
      <h2 align="center">The Report Card shows the detailed learning activities of learners who purchased content via ecommerce and meet the selection criteria below.</h2>
      </td>
    </tr>
    <tr>
      <td valign="top">
      <form method="POST" action="EcomReportCard.asp">
        <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th align="right" valign="top" width="30%">Select Start Date :</th>
            <td width="70%"><input type="text" name="vStrDate" size="15" value="<%=vStrDate%>"> <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br>
            ie Jan 1, 2008 (MMM DD, YYYY). Leave empty to start at first record.</td>
          </tr>
          <tr>
            <th align="right" valign="top" width="30%">End Date :</th>
            <td width="70%"><input type="text" name="vEndDate" size="15" value="<%=vEndDate%>"> <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br>
            ie Mar 31, 2008 (MMM DD, YYYY). Leave empty to finish with last record.</td>
          </tr>

          <tr>
            <th align="right" width="50%" valign="top">Find learners that :</th>
            <td width="50%"><input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>>start with<br><input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>>contain</td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">&nbsp;Learner ID :</td>
            <td width="50%"><input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">First Name :</td>
            <td width="50%"><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">Last Name :</td>
            <td width="50%"><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">Email Address :</td>
            <td width="50%"><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
          </tr>
          <tr>
            <th height="50" colspan="2"><input type="submit" value="Go" name="bGo" class="button"></th>
          </tr>
        </table>
      </form>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  <p><a href="EcomReportCard.asp">restart</a></p>

</body>

</html>


