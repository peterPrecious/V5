<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Phra.asp"-->
<!--#include virtual = "V5\Inc\Db_Mods.asp"-->

<%
	Dim vMods, aMods, vUrl_D, vUrl_S
	vMods = fDefault(Request("vMods"), "2676EN|2676FR|2677EN|2677FR|2678EN|2678FR|2699EN|2699FR|2719EN|2719FR|2800EN|2800FR|2830EN|2830FR|2835EN|2835FR|2843EN|2846EN|2846FR|2847EN|2847FR")
	aMods = Split(vMods, "|")
%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <script language="JavaScript" src="/V5/Inc/Launch.js"></script>
  <title>Vubiz Inactive Session</title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <div align="center">
    <table cellpadding="0" border="0" style="border-collapse: collapse" bordercolor="#111111" width="600">
      <tr>
        <th width="100%" align="left">
        <h1 align="center">Assessment Analysis Report</h1>
        <h2>Please click on either the Detail or the Summary buttons for the assessments listed below.</h2>
        <div align="center">
          <table border="1" width="90%" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#00FFFF">
            <tr>
              <th>Assessment ID </th>
              <th align="left">Title</th>
              <th colspan="2">Reports</th>
            </tr>
            <%
            	For i = 0 To Ubound(aMods)
            		sGetMods aMods(i)
            		If vMods_Type = "FX" Then
            			vUrl_D = "/Gold/vuClientReporting/ReportViewerFrame.aspx?AccountID=" & svCustAcctId & "&ModuleID=" & vMods_No & "&reportfile=App_Data/repMembAssmntDetails.frx"
            			vUrl_S = "/Gold/vuClientReporting/ReportViewerFrame.aspx?AccountID=" & svCustAcctId & "&ModuleID=" & vMods_No & "&reportfile=App_Data/repMembAssmntSummary.frx"
            %>            
            <tr>
              <td align="center"><%=vMods_Id%></td>
              <td><%=vMods_Title%></td>
              <td align="center"><input onclick="fullScreen('<%=vUrl_D%>')" type="button" value="Detail" name="bDetail" class="button85"></td>
              <td align="center"><input onclick="fullScreen('<%=vUrl_S%>')"  type="button" value="Summary" name="bSummary" class="button85"></td>
            </tr>
            <%
            		End If
            	Next
            %>
          </table>
          <p>Note: Multi-lingual Assessments Analysis is not yet available. <br>
          IE we currently do NOT integrate the results of modules 1234EN, 1234FR and 1234ES.</div>
        </th>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5\Inc\Shell_Lo.asp"-->

</body>

</html>
