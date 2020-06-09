<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vStrDate, vEndDate, vStrDateErr, vEndDateErr, vCustIdPrev, vNext
  
  '...default to previous month when you enter this site
  If Request.Form.Count = 0 Then

    vStrDateErr = "" : vStrDate = "Jan 1, 2000"
    vEndDateErr = "" : vEndDate = fFormatSqlDate(Now)

  Else

    vStrDate      = Trim(fFormatSqlDate(Request("vStrDate")))
    If vStrDate   = "" Then 
      vStrDate    = Request("vStrDate") '...put back bad date for display
      vStrDateErr = "Error"
    End If
    vEndDate      = Trim(fFormatSqlDate(Request("vEndDate")))
    If vEndDate   = "" Then
      vEndDate    = Request("vEndDate")  '...put back bad date for display
      vEndDateErr = "Error"
    End If
    If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
      If DateDiff("d", vStrDate, vEndDate) < 0 Then
        vEndDateErr = "Error"
      End If
    End If

    If vStrDateErr <> "Error" And vEndDateErr <> "Error" Then
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
  
      Response.Redirect fIf(Request("bOnline").Count > 0 , "SnapshotExam_O.asp", "SnapshotExam_X.asp") 
    End If
  End If
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Calendar.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <table border="1" width="600" cellpadding="0" style="border-collapse: collapse" bordercolor="#DDEEF9" cellspacing="0">
      <tr>
        <td valign="top" align="center" class="c2">
        <form method="POST" name="fDate" action="SnapshotExam.asp">
          <table border="0" cellspacing="0" cellpadding="2">
            <tr>
              <th valign="top" colspan="2" width="0" height="0"><h1 align="center"><br>Course Completion Snapshot</h1><p align="left">This report shows the number of Learners in this Parent account and all Child accounts (Facilitators are not included) who have Completed the course during the time frame selected based on passing the associated exam.<br>&nbsp;</p></th>
            </tr>
            <tr>
              <th align="right" valign="top" width="0" height="0">Select Start Date for Completion :</th>
              <td width="0" height="0">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                  <tr>
                    <td nowrap>
                      <input type="text" onblur="refillField('vStrDate', 'Jan 1, 2000')" name="vStrDate" id="vStrDate" size="12" value="<%=vStrDate%>" style="text-align: center" class="c2"> 
                      <a title="Start at the beginning..." class="debug" onclick="fillField('vStrDate', 'Jan 1, 2000')" href="#"><font color="#FFA500">&#937;</font></a> 
                      <a href="javascript:show_calendar('vStrDate','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>', 'MONTH DD YYYY');">
                      <img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a> 
                    </td>
                    <td align="right"><span style="background-color: #FFFF00"><%=vStrDateErr%></span></td>
                  </tr>
                  <tr>
                    <td colspan="2">ex Oct 1, <%=Year(Now - 1)%>(MMM D, YYYY) in English format</td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <th align="right" valign="top" width="0" height="0">End Date for Completion :</th>
              <td width="0" height="0">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                  <tr>
                    <td nowrap>
                      <input type="text" onblur="refillField('vEndDate', '<%=fFormatDate(Now)%>')" name="vEndDate" id="vEndDate" size="12" value="<%=vEndDate%>" style="text-align: center" class="c2"> 
                      <a title="End at today's date..." class="debug" onclick="fillField('vEndDate',  '<%=fFormatDate(DateAdd("d", 1, Now))%>')" href="#"><font color="#FFA500">&#937;</font></a> 
                      <a href="javascript:show_calendar('vEndDate','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>');"><img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a> 
                    </td>
                    <td align="right"><span style="background-color: #FFFF00"><%=vEndDateErr%></span> </td>
                  </tr>
                  <tr>
                    <td colspan="2">ex Dec 31, <%=Year(Now - 1)%>(MMM D, YYYY) in English format</td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td align="right" width="0" height="0">&nbsp;</td>
              <td align="right" width="0" height="0"><br>
                <input type="submit" value="Online" name="bOnline" class="button"> 
                <input type="submit" value="Excel"  name="bExcel"  class="button"> 
              </td>
            </tr>
          </table>
        </form>

        </td>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


