<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Phra.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Crit.asp"-->

<% 
  Dim vUrl, vDetails, vCurList, vStrDate, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFormat

  vDetails       = fDefault(Request("vDetails"), "n") 
  vCurList       = fDefault(Request("vCurList"), 0)
  vStrDate       = fDefault(Request("vStrDate"), fFormatSqlDate(DateAdd ("d", -90, Now)))
  vFind          = fDefault(Request("vFind"), "S")
  vFindId        = fUnQuote(Request("vFindId"))
  vFindFirstName = fUnQuote(Request("vFindFirstName"))
  vFindLastName  = fUnQuote(Request("vFindLastName"))
  vFindEmail     = fNoQuote(Request("vFindEmail"))
  vFormat        = fIf(Request("bExcel").Count = 1, "X", "O")

  '...processing the form?
  If Request("vForm").Count = 1 Then
    Session("soRs") = "" 
  
    '...goto online or excel reports
    vUrl = "Report" & vFormat & ".asp"     _
         & "?vStrDate="       & Server.UrlEncode(vStrDate)       _
         & "&vDetails="       & vDetails       _
         & "&vCurList="       & vCurList       _
         & "&vFind="          & vFind          _
         & "&vFindId="        & vFindId        _
         & "&vFindFirstName=" & vFindFirstName _
         & "&vFindLastName="  & vFindLastName  _
         & "&vFindEmail="     & vFindEmail
    Response.Redirect vUrl
'   Response.Write vUrl

  End If
  
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Launch.js"></script>
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td valign="top"><h1 align="center">Assessment | Survey Report </h1><h2>This displays learners who have taken an assessment or survey.&nbsp; You can either display the report online or create an <i>MS Excel</i> spreadsheet which can be saved on your <i>Desktop</i> 
		for analysis. <font color="#FF0000">Please be patient, the Excel format can take several minutes.</font> </h2></td>
    </tr>
    <tr>
      <td valign="top">
      <form method="POST" action="Report.asp">
        <table border="0" width="100%" cellpadding="2" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th align="right" nowrap valign="top" width="45%">Display :</th>
            <td width="55%"><input type="radio" value="y" name="vDetails" <%=fcheck("y", vdetails)%>>All Assessment attempts&nbsp;&nbsp; <br><input type="radio" value="n" name="vDetails" <%=fcheck("n", vdetails)%>>Highest Assessment grade achieved<br><input type="radio" value="s" name="vDetails" <%=fcheck("s", vdetails)%>>Surveys</td>
          </tr>
          <tr>
            <th align="right" nowrap valign="top" width="45%">
            <%    
              Dim vOption, vDesc, vSelected
              vOption = ""
              vSelected = ""
              For i = 1 To 9
                Select Case i
                  Case 1 : j =  1  : vDesc = "<!--{{-->1 day<!--}}-->"
                  Case 2 : j =  7  : vDesc = "7 " & "<!--{{-->days<!--}}-->"
                  Case 3 : j = 14  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 4 : j = 30  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 5 : j = 60  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 6 : j = 90  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 7 : j = 180 : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 8 : j = 365 : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 9 : j = 999 : vDesc = "<!--{{-->all available days<!--}}-->"
                End Select
                k = fFormatSqlDate(DateAdd ("d", -j, Now))
                vSelected = fIf(vStrDate = k, " selected", "")
                vOption = vOption & "<option value='" & k & "'" & vSelected & ">" & vDesc & "</option>" & vbCrLf 
              Next
            %> 
            Submitted during last :
            </th>
            <td width="55%"><select size="1" name="vStrDate"><%=vOption%></select></td>
          </tr>
          <tr>
            <th align="right" width="45%" valign="top">Selecting learners :</th>
            <td width="55%"><input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>>starting with <br><input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>>containing</td>
          </tr>
          <tr>
            <td align="right" width="45%" valign="top">&nbsp;Learner ID :</td>
            <td width="55%"><input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
          </tr>
          <tr>
            <td align="right" width="45%" valign="top">First Name :</td>
            <td width="55%"><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
          </tr>
          <tr>
            <td align="right" width="45%" valign="top">Last Name :</td>
            <td width="55%"><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
          </tr>
          <tr>
            <td align="right" width="45%" valign="top">Email Address :</td>
            <td width="55%"><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
          </tr>
          <tr>
            <th nowrap width="100%" height="50" colspan="2"><h2><br>Then click either....</h2><input type="submit" value="Display Online" name="bPrint" class="button"> or <input type="submit" value="MS Excel File" name="bExcel" class="button"><p>&nbsp;</th>
          </tr>
        </table>
        <input type="hidden" name="vForm" value="y">
      </form>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5\Inc\Shell_Lo.asp"-->

</body>

</html>
